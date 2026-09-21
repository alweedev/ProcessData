import numpy as np
import pandas as pd

from backend.domain.rules import FICHA_MAP, MODEL_COLS, REQUIRED_FICHA_FIELDS, SOURCE_ONLY_COLS
from backend.services.validation_service import ValidationService
from backend.shared.cpf_utils import clean_cpf, format_cpf_for_output
from backend.shared.excel_reader import read_first_sheet_as_text
from backend.shared.name_splitter import (
    CONFIDENT,
    DOUBTFUL,
    MAX_FIELD_LEN,
    get_vocabulary,
    split_full_name,
    suggest_abbreviation,
)
from backend.shared.text_utils import sanitize_output_text, upper_no_accents

# Colunas que identificam uma pessoa: linha com todas em branco é lixo da planilha (linha vazia), não registro.
_CRITICAL_COLS = ["Login", "NomeCompleto", "CPF", "Email"]

_BOOL_COLS = [
    "Solicitante",
    "Terceiro",
    "Vip",
    "ViajanteMasterNacional",
    "ViajanteMasterInternacional",
    "SolicitanteMaster",
    "MasterAdiantamento",
    "MasterReembolso",
]
_TRUE_SET = {"S", "SIM", "YES", "Y", "TRUE", "1"}

# Colunas de texto sanitizadas antes da validação (Nome e SobreNome não são cortados; a descrição do centro de
# custo só perde o acento, para preservar vírgulas, parênteses e barras: "COM AQUISICAO SFB (CO/N/NE), FASE 2").
_TEXT_COLS = [
    "Nome",
    "SobreNome",
    "NomeCompleto",
    "NomeEmpresa",
    "DescricaoCCustoEmpresa",
    "DescricaoCCustoCliente",
    "Cargo",
    "Departamento",
    "Cidade",
    "Estado",
    "Endereco",
]


def _map_unique(values, fn):
    """``[fn(v) for v in values]`` calculando cada valor distinto uma vez. As fichas repetem empresa, centro de
    custo, cargo, "S"/"N"... e as funções de limpeza são puras, então o resultado é idêntico e bem mais barato."""
    cache: dict = {}
    result = []
    for value in values:
        try:
            result.append(cache[value])
        except KeyError:
            result.append(cache.setdefault(value, fn(value)))
        except TypeError:  # valor não hasheável: calcula direto
            result.append(fn(value))
    return result


def _blank_row_mask(frame):
    critical = [c for c in _CRITICAL_COLS if c in frame.columns]
    if not critical:
        return pd.Series(False, index=frame.index)
    trimmed = frame[critical].apply(lambda s: s.astype(str).str.strip())
    return trimmed.eq("").all(axis=1)


def _row_label(file_number, excel_row, total_files):
    """Onde o usuário acha a linha: a ordem dos arquivos é a da lista que ele enviou."""
    return f"Arquivo {file_number} · linha {excel_row}" if total_files > 1 else f"Linha {excel_row}"


def _empty_file_problem(df, recognized_columns, sheet_name, sheet_count):
    """Por que uma planilha não rendeu nenhum registro (None se rendeu). Mensagens pensadas para o usuário
    conseguir corrigir a ficha: título acima do cabeçalho, dados em outra aba, coluna com nome fora do padrão."""
    if len(df.columns) == 0:
        problem = "A planilha está vazia."
    elif not recognized_columns:
        lidas = ", ".join(str(c)[:40] for c in list(df.columns)[:8])
        mais = "…" if len(df.columns) > 8 else ""
        problem = f"Nenhuma coluna da ficha foi reconhecida (colunas lidas: {lidas}{mais}). O cabeçalho precisa estar na 1ª linha."
    elif len(df) == 0:
        problem = "A planilha não tem linhas de dados abaixo do cabeçalho."
    else:
        return None
    if sheet_count > 1:
        problem += f" Só a 1ª aba ('{sheet_name}') é lida."
    return problem


def _map_bool_sn(value):
    # fold de acento antes de comparar: "Não"/"Sìm" normalizam
    return "S" if upper_no_accents(value).strip().upper() in _TRUE_SET else "N"


def _map_terceiro(value):
    # regra da ficha: se tiver dígitos, mantém os dígitos (id de terceiro);
    # senão mapeia Sim/Não -> S/N
    digits = "".join(filter(str.isdigit, str(value)))
    return digits if digits else _map_bool_sn(value)


def _prefix_front(value):
    if pd.isna(value) or str(value).strip() == "":
        return value
    return "FRONT" + str(value).replace(" ", "")


class ProcessingService:
    @staticmethod
    def _drop_header_like_rows(df):
        """Remove linhas que repetem o cabeçalho (mais de 40% das células iguais ao nome da coluna)."""
        if df.empty:
            return df
        columns = list(df.columns)
        matches = np.zeros(len(df), dtype=int)
        for position, column in enumerate(columns):
            same_as_header = df.iloc[:, position].astype(str).str.strip().str.upper() == str(column).upper()
            matches += same_as_header.to_numpy()
        return df[~((matches / max(1, len(columns))) > 0.4)]

    @staticmethod
    def process_records_from_files(paths, login_choice="CPF", fluxo="SELF", name_overrides=None):
        """Devolve ``(errors, df)``. ``df`` traz em ``df.attrs`` o que o relatório precisa e que
        não cabe nas colunas do arquivo de carga: ``duplicated_rows`` (linhas repetidas removidas),
        ``required_blank`` (em branco por campo obrigatório), ``row_labels`` e ``row_names`` (índice -> "Linha N" e
        o nome do passageiro, só das linhas com erro), ``name_review`` (nomes que pedem conferência: divisão duvidosa ou acima de 20
        caracteres) e ``name_overrides_applied`` (o que veio de ``name_overrides``, para o vocabulário aprender).

        ``name_overrides``: ``{"arquivo:linha": {"nome_completo", "nome", "sobrenome", "editado"}}`` — a decisão
        do usuário na conferência de nomes."""
        login_choice = str(login_choice or "CPF").strip().upper()
        normalized_map = {upper_no_accents(str(k)).strip(): v for k, v in FICHA_MAP.items()}
        parts = []  # um DataFrame por arquivo, já com os nomes de campo do sistema
        origin_file: list[int] = []  # nº do arquivo de cada registro
        origin_row: list[int] = []  # linha no Excel de cada registro
        all_errors = {}

        for file_number, path in enumerate(paths, start=1):
            try:
                if path.lower().endswith(".xlsx") or path.lower().endswith(".xls"):
                    df, sheet_name, sheet_count = read_first_sheet_as_text(path)
                    df = ProcessingService._drop_header_like_rows(df.fillna(""))
                    # campo do sistema -> coluna da ficha. Duas colunas para o mesmo campo: vale a última.
                    source_column = {}
                    for col in df.columns:
                        normalized_col = upper_no_accents(str(col)).strip()
                        if normalized_col in normalized_map:
                            source_column[normalized_map[normalized_col]] = col
                    problem = _empty_file_problem(df, source_column, sheet_name, sheet_count)
                    if problem:
                        all_errors[path] = problem
                    if source_column and len(df):
                        parts.append(pd.DataFrame({field: df[col].to_numpy() for field, col in source_column.items()}))
                        origin_file.extend([file_number] * len(df))
                        # +2: o cabeçalho é a linha 1 e o índice começa em 0
                        origin_row.extend((np.asarray(df.index, dtype=int) + 2).tolist())
            except Exception as exc:
                all_errors[path] = str(exc)

        if not parts:
            return all_errors, pd.DataFrame(columns=MODEL_COLS)

        # Com vários arquivos de colunas diferentes sobram células sem valor (NaN): viram vazio, senão
        # a validação não as enxerga como em branco e a limpeza de texto as grava como "NAN".
        df_final = pd.concat(parts, ignore_index=True).fillna("")
        # Antes de completar as colunas: só aqui dá para saber o que a ficha nem trazia. Uma coluna
        # ausente que tem fonte alternativa na ficha não conta como ausente.
        missing_required = [campo for campo in REQUIRED_FICHA_FIELDS if campo not in df_final.columns]
        alternative_sources = {"Telefone": "CelularContato", "Email": "EmailLogin"}
        if login_choice != "CPF":
            alternative_sources["CPF"] = "Passaporte"  # estrangeiro: o passaporte dispensa o CPF (por linha)
        for campo, fonte in alternative_sources.items():
            if fonte in df_final.columns and campo in missing_required:
                missing_required.remove(campo)
        for col in [*MODEL_COLS, *SOURCE_ONLY_COLS]:
            if col not in df_final.columns:
                df_final[col] = ""

        # Telefone em branco na ficha: o "CELULAR - CONTATO" é a segunda opção. Antes de validar, para o
        # telefone não ser acusado como vazio (nem a coluna, se só o contato foi preenchido).
        telefone_vazio = df_final["Telefone"].astype(str).str.strip().eq("")
        df_final.loc[telefone_vazio, "Telefone"] = df_final.loc[telefone_vazio, "CelularContato"]

        # "E-MAIL (LOGIN)" prevalece sobre o e-mail normal, que só vale quando o de login está em branco.
        usa_email_login = df_final["EmailLogin"].astype(str).str.strip().ne("")
        df_final.loc[usa_email_login, "Email"] = df_final.loc[usa_email_login, "EmailLogin"]

        df_final["Operacao"] = "INSERT"
        df_final["EmpresaCCustoParaUsuario"] = "S"
        df_final["CodigoIntegracao"] = "AUT"
        df_final["Status"] = ""

        # Nome (todos os nomes próprios) e SobreNome (todos os sobrenomes) saem do nome completo, como no documento.
        # Só sobrescreve o que ele deu: sem nome completo, valem o NOME e o SOBRENOME da ficha.
        vocabulary = get_vocabulary()
        splits = _map_unique(df_final["NomeCompleto"].tolist(), lambda v: split_full_name(v, vocabulary))
        first_names = df_final["Nome"].tolist()
        last_names = df_final["SobreNome"].tolist()
        for position, split in enumerate(splits):
            if split.nome:
                first_names[position] = split.nome
            if split.sobrenome:
                last_names[position] = split.sobrenome

        # O que o usuário confirmou ou corrigiu na tela de conferir vale mais que a sugestão. A chave é a linha da ficha
        # ("arquivo:linha") e o override só vale se o nome completo ainda for o mesmo: planilha trocada não herda nada.
        applied_overrides: dict[int, dict] = {}
        for position in range(len(df_final)) if name_overrides else ():
            override = name_overrides.get(f"{origin_file[position]}:{origin_row[position]}")
            if not override:
                continue
            if sanitize_output_text(df_final["NomeCompleto"].iat[position], None) != override["nome_completo"]:
                continue
            first_names[position] = sanitize_output_text(override["nome"], None)
            last_names[position] = sanitize_output_text(override["sobrenome"], None)
            applied_overrides[position] = override
        df_final["Nome"] = first_names
        df_final["SobreNome"] = last_names

        if login_choice == "CPF":
            if "CPF" in df_final.columns:
                df_final["Login"] = _map_unique(
                    df_final["CPF"].tolist(), lambda x: format_cpf_for_output(clean_cpf(x)) if x else ""
                )
        elif login_choice == "EMAIL":
            if "Email" in df_final.columns:
                df_final["Login"] = df_final["Email"]

        fluxo_up = (fluxo or "").upper()
        if fluxo_up == "SELF":
            for col in [
                "Vip",
                "ViajanteMasterNacional",
                "ViajanteMasterInternacional",
                "SolicitanteMaster",
                "MasterAdiantamento",
                "MasterReembolso",
            ]:
                df_final[col] = "N"
        elif fluxo_up == "FRONT":
            df_final["ViajanteMasterNacional"] = "S"
            df_final["ViajanteMasterInternacional"] = "S"
            for col in ["Vip", "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]:
                df_final[col] = "N"
            if "Login" in df_final.columns:
                df_final["Login"] = _map_unique(df_final["Login"].tolist(), _prefix_front)

        for col in _TEXT_COLS:
            if col in df_final.columns:
                # Nome e SobreNome NÃO são mais cortados em 20: cortar mudaria a identidade do passageiro em silêncio.
                # O que passar do limite vai para a conferência (`name_review`) e a geração espera o ajuste.
                if col == "DescricaoCCustoEmpresa":
                    df_final[col] = _map_unique(df_final[col].tolist(), upper_no_accents)
                else:
                    df_final[col] = _map_unique(df_final[col].tolist(), lambda v: sanitize_output_text(v, None))

        errors = dict(all_errors)  # arquivos que não puderam ser lidos seguem no resultado, junto dos erros por linha
        records = df_final.to_dict("records")
        for position, record in enumerate(records):
            msgs = ValidationService.validate_row(record, login_choice)
            if msgs:
                errors[df_final.index[position]] = "; ".join(msgs)
        if "Nivel" in df_final.columns:  # a validação autocorrige o Nível no registro; leva a correção para o resultado
            df_final["Nivel"] = [record["Nivel"] for record in records]

        geral = ValidationService.validate_dataframe(df_final, missing_required, login_choice)
        if geral:
            errors["__geral__"] = "; ".join(geral)

        duplicated_rows = 0
        if "Login" in df_final.columns and "NomeCompleto" in df_final.columns:
            key = ["Login", "NomeCompleto"]
            # Conta antes de remover (depois não sobra o que contar) e só entre linhas com dados:
            # linhas em branco iguais entre si não são "duplicadas".
            duplicated_rows = int(df_final[~_blank_row_mask(df_final)].duplicated(subset=key).sum())
            df_final = df_final.drop_duplicates(subset=key, keep="first")

        for bc in _BOOL_COLS:
            if bc not in df_final.columns:
                df_final[bc] = "N"
            else:
                fn = _map_terceiro if bc == "Terceiro" else _map_bool_sn
                df_final[bc] = _map_unique(df_final[bc].fillna("").tolist(), fn)

        if "NroMatricula" in df_final.columns:
            df_final["NroMatricula"] = _map_unique(
                df_final["NroMatricula"].fillna("").tolist(), lambda v: "".join(filter(str.isdigit, str(v)))
            )

        for col in df_final.columns:
            if df_final[col].dtype == object:
                values = df_final[col].fillna("").astype(str).tolist()
                if col in ("Email", "Telefone") or (col == "Login" and login_choice == "EMAIL"):
                    df_final[col] = _map_unique(values, lambda v: v.strip().upper())
                elif col == "DescricaoCCustoEmpresa":
                    df_final[col] = _map_unique(values, upper_no_accents)
                else:
                    df_final[col] = _map_unique(values, lambda v: sanitize_output_text(v, None))

        df_final = df_final.loc[~_blank_row_mask(df_final)].copy()

        # Os números do relatório valem para o que de fato sai no arquivo: erro de linha descartada
        # (repetida ou em branco) não conta como linha inválida.
        errors = {k: v for k, v in errors.items() if isinstance(k, str) or k in df_final.index}
        required_blank = {}
        for campo, rotulo in REQUIRED_FICHA_FIELDS.items():
            blank = df_final[campo].astype(str).str.strip().eq("")
            if campo == "CPF" and login_choice != "CPF":
                blank &= df_final["Passaporte"].astype(str).str.strip().eq("")  # com passaporte, CPF em branco é ok
            required_blank[rotulo] = int(blank.sum())
        # Só as linhas com erro precisam de rótulo. `attrs` é copiado em profundidade a cada operação do
        # pandas: um rótulo por linha (20 mil) chegava a custar um terço do tempo de exportar o arquivo.
        total_files = len(paths)
        error_rows = [idx for idx in errors if not isinstance(idx, str)]
        row_labels = {idx: _row_label(origin_file[idx], origin_row[idx], total_files) for idx in error_rows}
        # Quem é o passageiro de cada linha com problema: achar a linha na ficha só pelo número é trabalhoso.
        row_names = {idx: df_final.at[idx, "NomeCompleto"] for idx in error_rows}

        # Nomes que pedem conferência: divisão duvidosa (e ainda não decidida pelo usuário) ou acima do limite da
        # plataforma. Só esses vão para `attrs` (poucos): a lista inteira custaria caro, como os rótulos acima.
        has_name = df_final["NomeCompleto"].astype(str).str.strip().ne("")
        doubtful = np.array(
            [splits[idx].confidence == DOUBTFUL and idx not in applied_overrides for idx in df_final.index]
        )
        too_long_first = df_final["Nome"].astype(str).str.len() > MAX_FIELD_LEN
        too_long_last = df_final["SobreNome"].astype(str).str.len() > MAX_FIELD_LEN
        needs_review = (has_name.to_numpy() & doubtful) | too_long_first.to_numpy() | too_long_last.to_numpy()
        name_review = []
        for idx in df_final.index[needs_review]:
            first, last = df_final.at[idx, "Nome"], df_final.at[idx, "SobreNome"]
            long_first, long_last = len(first) > MAX_FIELD_LEN, len(last) > MAX_FIELD_LEN
            name_review.append(
                {
                    "key": f"{origin_file[idx]}:{origin_row[idx]}",
                    "label": _row_label(origin_file[idx], origin_row[idx], total_files),
                    "nome_completo": df_final.at[idx, "NomeCompleto"],
                    "nome": first,
                    "sobrenome": last,
                    "confianca": DOUBTFUL if doubtful[df_final.index.get_loc(idx)] else CONFIDENT,
                    "motivos": list(splits[idx].reasons) if doubtful[df_final.index.get_loc(idx)] else [],
                    "estouro": {"nome": long_first, "sobrenome": long_last},
                    "sugestao": (
                        dict(zip(("nome", "sobrenome"), suggest_abbreviation(first, last), strict=True))
                        if long_first or long_last
                        else None
                    ),
                }
            )
        overrides_applied = [
            {"nome": o["nome"], "sobrenome": o["sobrenome"], "editado": bool(o.get("editado"))}
            for position, o in applied_overrides.items()
            if position in df_final.index
        ]

        result = df_final[MODEL_COLS]
        result.attrs.update(
            duplicated_rows=duplicated_rows,
            required_blank=required_blank,
            row_labels=row_labels,
            row_names=row_names,
            name_review=name_review,
            name_overrides_applied=overrides_applied,
        )
        return errors, result
