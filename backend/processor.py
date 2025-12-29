import os
import re
import logging
import pandas as pd
from docx import Document
from .utils import upper_no_accents, limpar_cpf_raw, format_cpf_for_output, separar_nome_sobrenome
from .validators import validar_linha


# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("robo_backend")

MODEL_COLS = [
    "Operacao", "UserId", "Login", "CodigoCCustoCliente", "DescricaoCCustoCliente",
    "NomeEmpresa", "CodigoCCustoEmpresa", "DescricaoCCustoEmpresa", "EmpresaCCustoParaUsuario",
    "NroMatricula", "Nome", "SobreNome", "NomeCompleto", "Email", "Telefone", "Cargo", "Departamento", "Nivel",
    "Endereco", "Cidade", "Estado", "CEP", "Solicitante", "Vip", "ViajanteMasterNacional",
    "ViajanteMasterInternacional", "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso", "Terceiro",
    "CodigoIntegracao", "Status"
]


# Helpers de sanitização e extração usados em cadastro/inativação
def extract_digits_only(v: str) -> str:
    """Retorna apenas os dígitos da string (ou '' se nenhum dígito)."""
    try:
        s = str(v)
    except Exception:
        return ""
    digits = re.sub(r"\D", "", s)
    return digits


def sanitize_output_text(v: str, maxlen: int | None = None) -> str:
    """Remove acentos e caracteres especiais, mantém letras, números e espaços.
    Retorna em MAIÚSCULAS. Opcionalmente trunca para maxlen.
    """
    if v is None:
        return ""
    s = str(v)
    s = upper_no_accents(s)
    # manter apenas letras, dígitos e espaços
    s = re.sub(r"[^A-Z0-9 \-]", "", s.upper())
    s = re.sub(r"\s+", " ", s).strip()
    if maxlen:
        return s[:maxlen]
    return s


def split_name_first_last(fullname: str) -> tuple:
    """Separa apenas primeiro e último nome a partir de `NomeCompleto`.
    Regra simples: primeiro token como Nome e último token como SobreNome.
    Não tenta detectar partículas (da/de/dos etc.) nem sobrenomes compostos.
    """
    if not fullname:
        return "", ""

    parts = [p for p in str(fullname).strip().split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        first, last = parts[0], ""
    else:
        first, last = parts[0], parts[-1]

    # Sanitizar e aplicar limites de 20 caracteres
    first_clean = sanitize_output_text(first, 20)
    last_clean = sanitize_output_text(last, 20)

    return first_clean, last_clean


FICHA_MAP = {
    "CPF": "CPF",
    "CPF (SEM PONTOS)": "CPF",
    "EMPRESA (DO GRUPO)": "NomeEmpresa",
    "Empresa": "NomeEmpresa",
    "Centro de custo": "CodigoCCustoEmpresa",
    # Suportes com underscore conforme pedido
    "Centro_de_Custo": "CodigoCCustoEmpresa",
    "CODIGO - CENTRO DE CUSTO": "CodigoCCustoEmpresa",
    "DESCRICAO - CENTRO DE CUSTO": "DescricaoCCustoEmpresa",
    "Descrição Centro de Custo": "DescricaoCCustoEmpresa",
    # Suportes com underscore conforme pedido
    "Codigo_Centro_De_Custo": "DescricaoCCustoEmpresa",
    "MATRICULA": "NroMatricula",
    "Matricula": "NroMatricula",
    "NroMatricula": "NroMatricula",
    "NOME": "Nome",
    "SOBRENOME (ATE 20 CARACTERES)": "SobreNome",
    "NOME COMPLETO": "NomeCompleto",
    "NomeCompleto": "NomeCompleto",
    "EMAIL": "Email",
    "E-MAIL": "Email",
    "TELEFONE": "Telefone",
    "CARGO": "Cargo",
    "DEPARTAMENTO": "Departamento",
    "NIVEL": "Nivel",
    "NÍVEL": "Nivel",
    "SOLICITANTE? (S/N)": "Solicitante",
    "TERCEIRO? (S/N)": "Terceiro",
    "Terceiro": "Terceiro",
    "ENVIA DADOS DE ACESSO? (S/N)": "EnviaDadosAcesso",
}


def drop_header_like_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas que parecem ser cabeçalhos repetidos dentro do arquivo Excel."""
    if df.empty:
        return df
    cols = list(df.columns)

    def is_header(row):
        matches = 0
        for c in cols:
            val = str(row.get(c, "")).strip()
            if val.upper() == str(c).upper():
                matches += 1
        return (matches / max(1, len(cols))) > 0.4

    return df[~df.apply(is_header, axis=1)]


def extrair_docx(path: str) -> dict:
    """Extrai pares label:value de um .docx usando FICHA_MAP como referência."""
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    data = {}
    for label, target in FICHA_MAP.items():
        pat = rf"{re.escape(label)}\s*[:\-]\s*(.+)"
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            data[target] = m.group(1).strip()
    return data


def processar_registros_from_files(paths: list, login_choice: str = "CPF", fluxo: str = "SELF"):
    """Processa arquivos (.docx, .xls, .xlsx) e retorna (errors, df_final)."""
    all_errors = {}
    all_data = []

    for path in paths:
        try:
            if path.lower().endswith('.docx'):
                data = extrair_docx(path)
                if data:
                    all_data.append(data)
            elif path.lower().endswith(('.xls', '.xlsx')):
                df = pd.read_excel(path, dtype=str).fillna("")
                df = drop_header_like_rows(df)
                normalized_map = {upper_no_accents(k).strip(): v for k, v in FICHA_MAP.items()}
                for _, row in df.iterrows():
                    mapped_row = {}
                    for col in df.columns:
                        normalized_col = upper_no_accents(str(col)).strip()
                        if normalized_col in normalized_map:
                            mapped_row[normalized_map[normalized_col]] = row[col]
                    if mapped_row:
                        all_data.append(mapped_row)
            else:
                logger.debug(f"Ignorando arquivo não suportado: {path}")
        except Exception as e:
            logger.warning(f"Falha ao ler {path}: {e}")
            all_errors[path] = str(e)

    if not all_data:
        return all_errors, pd.DataFrame(columns=MODEL_COLS)

    df_final = pd.DataFrame(all_data)

    for col in MODEL_COLS:
        if col not in df_final.columns:
            df_final[col] = ""

    df_final["Operacao"] = "INSERT"
    df_final["EmpresaCCustoParaUsuario"] = "S"
    df_final["CodigoIntegracao"] = "AUT"
    df_final["Status"] = ""

    for idx, row in df_final.iterrows():
        # Sempre recalcular Nome e SobreNome a partir de NomeCompleto,
        # dando prioridade à lógica do script em relação ao que veio na ficha.
        first, last = split_name_first_last(row.get("NomeCompleto", ""))
        if first:
            df_final.at[idx, "Nome"] = first
        if last:
            df_final.at[idx, "SobreNome"] = last

    if login_choice == "CPF":
        if "CPF" in df_final.columns:
            df_final["Login"] = df_final["CPF"].apply(
                lambda x: format_cpf_for_output(limpar_cpf_raw(x)) if x else ""
            )
    elif login_choice == "EMAIL":
        if "Email" in df_final.columns:
            df_final["Login"] = df_final["Email"]

    try:
        fluxo_up = (fluxo or "").upper()
    except Exception:
        fluxo_up = ""
    if fluxo_up == "SELF":
        for col in ["Vip", "ViajanteMasterNacional", "ViajanteMasterInternacional",
                    "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]:
            df_final[col] = "N"
    elif fluxo_up == "FRONT":
        df_final["ViajanteMasterNacional"] = "S"
        df_final["ViajanteMasterInternacional"] = "S"
        for col in ["Vip", "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]:
            df_final[col] = "N"
        if "Login" in df_final.columns:
            def prefix_front(v):
                if pd.isna(v) or str(v).strip() == "":
                    return v
                s = str(v)
                return "FRONT" + s.replace(" ", "")
            df_final["Login"] = df_final["Login"].apply(prefix_front)

    text_cols = [
        "Nome", "SobreNome", "NomeCompleto", "NomeEmpresa",
        "DescricaoCCustoEmpresa", "DescricaoCCustoCliente", "Cargo",
        "Departamento", "Cidade", "Estado", "Endereco"
    ]
    for c in text_cols:
        if c in df_final.columns:
            if c == 'Nome':
                df_final[c] = df_final[c].apply(lambda v: sanitize_output_text(v, 20))
            elif c == 'SobreNome':
                df_final[c] = df_final[c].apply(lambda v: sanitize_output_text(v, 20))
            elif c == 'NomeCompleto':
                df_final[c] = df_final[c].apply(lambda v: sanitize_output_text(v, None))
            else:
                df_final[c] = df_final[c].apply(lambda v: sanitize_output_text(v, None))

    errors = {}
    for idx, row in df_final.iterrows():
        msgs = validar_linha(row)
        if msgs:
            errors[idx] = "; ".join(msgs)

    if "Login" in df_final.columns and "NomeCompleto" in df_final.columns:
        df_final = df_final.drop_duplicates(subset=["Login", "NomeCompleto"], keep="first")

    # Normalizar campos booleanos (mapear Sim/Não, Yes/No, True/False para S/N)
    # Garantir que 'Solicitante' exista e seja preenchido (obrigatório na saída)
    def map_bool_to_SN(v):
        try:
            s = upper_no_accents(str(v)).strip().upper()
        except Exception:
            s = str(v).strip().upper()
        if s in ("S", "SIM", "YES", "Y", "TRUE", "1"):
            return "S"
        return "N"

    bool_cols = ["Solicitante", "Terceiro", "Vip", "ViajanteMasterNacional", "ViajanteMasterInternacional",
                 "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]
    for bc in bool_cols:
        if bc not in df_final.columns:
            # 'Solicitante' é obrigatório; outros campos recebem 'N' por padrão
            df_final[bc] = "N"
        else:
            if bc == 'Terceiro':
                # se houver dígitos, manter apenas os dígitos; caso contrário, mapear Sim/Não para S/N
                df_final[bc] = df_final[bc].fillna("").apply(lambda v: extract_digits_only(v) if extract_digits_only(v) else map_bool_to_SN(v))
            else:
                df_final[bc] = df_final[bc].fillna("").apply(map_bool_to_SN)

    # Garantir que NroMatricula contenha apenas dígitos
    if 'NroMatricula' in df_final.columns:
        df_final['NroMatricula'] = df_final['NroMatricula'].fillna('').apply(lambda v: extract_digits_only(v))

    # Sanitizar restante das colunas de texto para saída (uppercase, sem acentos, colapso de espaços)
    # EXCEÇÃO: manter acentos e caracteres especiais em 'Email' e 'Telefone' (apenas trim + uppercase)
    for col in df_final.columns:
        if df_final[col].dtype == object:
            if col in ("Email", "Telefone"):
                df_final[col] = df_final[col].fillna('').astype(str).apply(lambda v: v.strip().upper())
            elif col == 'NroMatricula':
                df_final[col] = df_final[col].fillna('').apply(lambda v: extract_digits_only(v))
            else:
                df_final[col] = df_final[col].fillna('').astype(str).apply(lambda v: sanitize_output_text(v, None))

    # Descartar linhas em branco (apenas espaços) sem dados críticos
    def _drop_blank_rows(df: pd.DataFrame) -> pd.DataFrame:
        critical = [c for c in ["Login", "NomeCompleto", "CPF", "Email"] if c in df.columns]
        if not critical:
            return df
        trimmed = df[critical].apply(lambda s: s.astype(str).str.strip())
        mask_blank = trimmed.eq("").all(axis=1)
        return df.loc[~mask_blank].copy()

    df_final = _drop_blank_rows(df_final)

    df_final = df_final[MODEL_COLS]

    return errors, df_final


# ==========================================================
# NOVA VERSÃO: processar_inativacao_from_paths (compatível)
# ==========================================================
def processar_inativacao_from_paths(df_base: pd.DataFrame, df_lista: pd.DataFrame,
                                    use_fuzzy: bool = False, fuzzy_cutoff: float = 0.9):
    """
    Processa inativação comparando usuários da base com uma lista de desligados.
    Estratégia:
      - Match exato por CPF (prioritário)
      - Match exato por NomeCompleto (fallback)
    Retorna: (df_inativacao, stats)
    """
    try:
        def normalize_str(s):
            return upper_no_accents(str(s)).strip() if pd.notna(s) else ""

        def normalize_cpf(s):
            s = re.sub(r"\D", "", str(s))
            return s.zfill(11) if s else ""

        df_base = df_base.copy()
        df_lista = df_lista.copy()

        # Detectar colunas relevantes
        col_map = {upper_no_accents(str(c)).strip(): c for c in df_base.columns}
        cpf_col = next((v for k, v in col_map.items() if "CPF" in k), None)
        logger.info("Coluna CPF detectada: {}".format(cpf_col) if cpf_col else "Nenhuma coluna CPF detectada na base; CPF matching desabilitado")
        nome_col = next((v for k, v in col_map.items() if "NOMECOMPLETO" in k or "NOME COMPLETO" in k), None)
        email_col = next((v for k, v in col_map.items() if "EMAIL" in k), None)
        status_col = next((v for k, v in col_map.items() if "STATUS" in k), None)

        df_base["CPFdigits"] = df_base[cpf_col].apply(normalize_cpf) if cpf_col else ""
        df_base["Nome Normalizado"] = df_base[nome_col].apply(normalize_str) if nome_col else ""
        df_base["Email Normalizado"] = df_base[email_col].astype(str).fillna("").str.strip().str.lower() if email_col else ""

        if status_col:
            df_base["Status Normalizado"] = df_base[status_col].apply(normalize_str)
            df_base = df_base[df_base["Status Normalizado"] == "ATIVO"].copy()

        df_lista["CPFdigits"] = df_lista["CPF"].apply(normalize_cpf) if "CPF" in df_lista.columns else ""
        df_lista["Nome Normalizado"] = df_lista["NomeCompleto"].apply(normalize_str) if "NomeCompleto" in df_lista.columns else ""
        df_lista["Email Normalizado"] = df_lista["Email"].astype(str).fillna("").str.strip().str.lower() if "Email" in df_lista.columns else ""

        lista_cpfs = [cpf for cpf in df_lista["CPFdigits"].unique() if cpf]
        matched_by_cpf = df_base[df_base["CPFdigits"].isin(lista_cpfs)] if cpf_col else pd.DataFrame()

        lista_nomes = [nome for nome in df_lista["Nome Normalizado"].unique() if nome]
        matched_by_nome = df_base[
            (df_base["Nome Normalizado"].isin(lista_nomes)) &
            (~df_base.index.isin(matched_by_cpf.index))
        ] if nome_col else pd.DataFrame()

        lista_emails = [em for em in df_lista["Email Normalizado"].unique() if em]
        matched_by_email = df_base[
            (df_base["Email Normalizado"].isin(lista_emails)) &
            (~df_base.index.isin(matched_by_cpf.index)) &
            (~df_base.index.isin(matched_by_nome.index))
        ] if email_col else pd.DataFrame()

        stats = {
            "cpf_matches": len(matched_by_cpf),
            "name_matches": len(matched_by_nome),
            "email_matches": len(matched_by_email)
        }

        # adicionar coluna temporária de match_type para auditoria
        if not matched_by_cpf.empty:
            matched_by_cpf = matched_by_cpf.copy()
            matched_by_cpf["__match_type"] = "cpf"
        if not matched_by_nome.empty:
            matched_by_nome = matched_by_nome.copy()
            matched_by_nome["__match_type"] = "nome"
        if not matched_by_email.empty:
            matched_by_email = matched_by_email.copy()
            matched_by_email["__match_type"] = "email"

        matched = pd.concat([matched_by_cpf, matched_by_nome, matched_by_email], ignore_index=True).drop_duplicates()
        # normalizar índices para evitar problemas ao extrair colunas por posição
        matched = matched.reset_index(drop=True)

        if matched.empty:
            logger.warning("Nenhuma correspondência encontrada para inativação.")
            # garantir que stats contenha mapeamento vazio de inactive_matches
            stats["inactive_matches"] = {}
            return pd.DataFrame(columns=MODEL_COLS), stats

        out_df = pd.DataFrame(index=range(len(matched)), columns=MODEL_COLS)
        out_df["Operacao"] = "DELETE"

        def pick(df, *keys):
            for k in keys:
                if k in df.columns:
                    return df[k].values
            return [""] * len(df)

        # Robust lookup by alias names (ignores spaces, underscores, case, and accents)
        def _norm_key(s: str) -> str:
            try:
                return re.sub(r"[^A-Z0-9]", "", upper_no_accents(str(s)).upper())
            except Exception:
                return ""

        def pick_by_alias(df, *aliases):
            if df is None or df.empty:
                return [""] * (0 if df is None else len(df))
            alias_norms = [_norm_key(a) for a in aliases]
            for col in df.columns:
                nk = _norm_key(col)
                for an in alias_norms:
                    if an and (nk == an or an in nk):
                        return df[col].values
            return [""] * len(df)

        out_df["UserId"] = pick(matched, "UserId")
        out_df["Login"] = pick(matched, "Login", "UserName")
        out_df["NomeCompleto"] = pick(matched, "NomeCompleto", "Nome Completo", nome_col)
        out_df["Nome"] = pick(matched, "Nome")
        out_df["SobreNome"] = pick(matched, "SobreNome")
        out_df["Email"] = pick(matched, "Email")
        out_df["Telefone"] = pick(matched, "Telefone")
        out_df["Cargo"] = pick(matched, "Cargo")
        out_df["Departamento"] = pick(matched, "Departamento")
        out_df["Nivel"] = pick(matched, "Nivel")
        out_df["NomeEmpresa"] = pick(matched, "Empresa")
        #busca a empresa, centro de custo e descrição que estiver configurado no usuário.
        out_df["CodigoCCustoEmpresa"] = pick(matched, "Codigo_Centro_de_Custo" , )
        out_df["DescricaoCCustoEmpresa"] = pick(matched, "Centro_de_Custo")
        out_df["EmpresaCCustoParaUsuario"] = "S"
        out_df["CodigoIntegracao"] = "AUT"
        out_df["Status"] = ""

        # Valores padrão para campos que serão preenchidos com mapeamento booleano
        bool_defaults = [
            "Solicitante", "Vip", "ViajanteMasterNacional", "ViajanteMasterInternacional",
            "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso", "Terceiro"
        ]
        for c in ["Endereco", "Cidade", "Estado", "CEP"]:
            out_df[c] = ""
        for c in bool_defaults:
            out_df[c] = "N"

        # Preencher flags booleanas a partir das colunas reais na base (detecção robusta)
        def find_real_col(target_name):
            # procura por uma coluna cujo nome normalizado contenha target_name (ex: 'SOLICITANTE')
            t = upper_no_accents(target_name).strip()
            for k, v in col_map.items():
                if t in k:
                    return v
            return None

        bool_map = {"SIM": "S", "NAO": "N", "NÃO": "N", "S": "S", "N": "N", "TRUE": "S", "FALSE": "N"}
        for logical_col in bool_defaults:
            real_col = find_real_col(logical_col)
            if real_col and real_col in matched.columns:
                # Para 'Terceiro', priorizar dígitos se existirem; caso contrário mapear Sim/Não para S/N
                if logical_col == 'Terceiro':
                    def terceiro_map(v):
                        d = extract_digits_only(v)
                        if d:
                            return d
                        return bool_map.get(str(v).strip().upper(), "N")
                    vals = matched[real_col].fillna("").astype(str).map(lambda x: terceiro_map(x))
                    out_df[logical_col] = vals.values
                else:
                    # normalizar valores e mapear para S/N
                    vals = matched[real_col].fillna("").astype(str).str.strip().str.upper().map(lambda x: bool_map.get(x, "N"))
                    out_df[logical_col] = vals.values

        # Para inativação, manter Nome e SobreNome exatamente como estão na base
        # (não recalcular a partir de NomeCompleto), apenas garantir que colunas existam.
        if "Nome" not in out_df.columns:
            out_df["Nome"] = ""
        if "SobreNome" not in out_df.columns:
            out_df["SobreNome"] = ""

        # Garantir NroMatricula preenchido a partir de 'Matricula' caso necessário e somente com dígitos
        try:
            nro_vals = pick_by_alias(matched, "NroMatricula", "Matricula")
            out_df["NroMatricula"] = [extract_digits_only(v) for v in nro_vals]
        except Exception:
            out_df["NroMatricula"] = ""

        out_df = out_df[MODEL_COLS]

        # Padronizar: todos os campos em MAIÚSCULAS na ficha de saída (inativação)
        for col in out_df.columns:
            if out_df[col].dtype == object:
                # garantir string, remover espaços nas bordas e aplicar upper; dígitos permanecem inalterados
                out_df[col] = out_df[col].fillna("").astype(str).str.strip().str.upper()

        # Construir mapeamento de inactive_matches para o preview (listas de dicionários)
        inactive = {}
        try:
            if not matched_by_cpf.empty:
                inactive['cpf'] = matched_by_cpf.fillna('').to_dict(orient='records')
            else:
                inactive['cpf'] = []
        except Exception:
            inactive['cpf'] = []
        try:
            if not matched_by_nome.empty:
                inactive['nome'] = matched_by_nome.fillna('').to_dict(orient='records')
            else:
                inactive['nome'] = []
        except Exception:
            inactive['nome'] = []
        try:
            if not matched_by_email.empty:
                inactive['email'] = matched_by_email.fillna('').to_dict(orient='records')
            else:
                inactive['email'] = []
        except Exception:
            inactive['email'] = []

        stats['inactive_matches'] = inactive
        # total de matches combinados (fonte de verdade para contagem no preview)
        try:
            stats['total_matches'] = int(matched.shape[0])
        except Exception:
            stats['total_matches'] = sum(len(v) for v in inactive.values() if isinstance(v, list))

        logger.info(
            f"Inativação concluída. Linhas encontradas: {out_df.shape[0]} (CPF={stats['cpf_matches']}, Nome={stats['name_matches']}, Email={stats['email_matches']})"
        )

        return out_df, stats

    except Exception as e:
        logger.error(f"Erro em processar_inativacao_from_paths: {e}")
        # garantir que stats sempre tenha total_matches válido mesmo em caso de erro
        error_stats = {"error": str(e), "cpf_matches": 0, "name_matches": 0, "total_matches": 0, "inactive_matches": {}}
        return pd.DataFrame(columns=MODEL_COLS), error_stats
    # Verificar colunas obrigatórias