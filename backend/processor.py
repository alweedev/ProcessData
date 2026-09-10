import re
from typing import Any

import pandas as pd

from .core.logging import get_logger
from .domain.rules import MODEL_COLS
from .shared.text_utils import upper_no_accents

logger = get_logger()


# Helper de extração usado por processar_inativacao_from_paths.
# O pipeline de cadastro foi consolidado em backend.services.processing_service.
def extract_digits_only(v: str) -> str:
    """Retorna apenas os dígitos da string (ou '' se nenhum dígito)."""
    try:
        s = str(v)
    except Exception:
        return ""
    return re.sub(r"\D", "", s)


# ==========================================================
# processar_inativacao_from_paths
# ==========================================================
def processar_inativacao_from_paths(df_base: pd.DataFrame, df_lista: pd.DataFrame):
    """
    Processa inativação comparando usuários da base com uma lista de desligados.
    Estratégia:
      - Match exato por CPF (prioritário)
      - Match exato por NomeCompleto (fallback)
      - Match exato por Email (fallback)
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
        logger.info(
            f"Coluna CPF detectada: {cpf_col}"
            if cpf_col
            else "Nenhuma coluna CPF detectada na base; CPF matching desabilitado"
        )
        nome_col = next((v for k, v in col_map.items() if "NOMECOMPLETO" in k or "NOME COMPLETO" in k), None)
        email_col = next((v for k, v in col_map.items() if "EMAIL" in k), None)
        status_col = next((v for k, v in col_map.items() if "STATUS" in k), None)

        df_base["CPFdigits"] = df_base[cpf_col].apply(normalize_cpf) if cpf_col else ""
        df_base["Nome Normalizado"] = df_base[nome_col].apply(normalize_str) if nome_col else ""
        df_base["Email Normalizado"] = (
            df_base[email_col].astype(str).fillna("").str.strip().str.lower() if email_col else ""
        )

        if status_col:
            df_base["Status Normalizado"] = df_base[status_col].apply(normalize_str)
            df_base = df_base[df_base["Status Normalizado"] == "ATIVO"].copy()

        df_lista["CPFdigits"] = df_lista["CPF"].apply(normalize_cpf) if "CPF" in df_lista.columns else ""
        df_lista["Nome Normalizado"] = (
            df_lista["NomeCompleto"].apply(normalize_str) if "NomeCompleto" in df_lista.columns else ""
        )
        df_lista["Email Normalizado"] = (
            df_lista["Email"].astype(str).fillna("").str.strip().str.lower() if "Email" in df_lista.columns else ""
        )

        lista_cpfs = [cpf for cpf in df_lista["CPFdigits"].unique() if cpf]
        matched_by_cpf = df_base[df_base["CPFdigits"].isin(lista_cpfs)] if cpf_col else pd.DataFrame()

        lista_nomes = [nome for nome in df_lista["Nome Normalizado"].unique() if nome]
        matched_by_nome = (
            df_base[(df_base["Nome Normalizado"].isin(lista_nomes)) & (~df_base.index.isin(matched_by_cpf.index))]
            if nome_col
            else pd.DataFrame()
        )

        lista_emails = [em for em in df_lista["Email Normalizado"].unique() if em]
        matched_by_email = (
            df_base[
                (df_base["Email Normalizado"].isin(lista_emails))
                & (~df_base.index.isin(matched_by_cpf.index))
                & (~df_base.index.isin(matched_by_nome.index))
            ]
            if email_col
            else pd.DataFrame()
        )

        stats: dict[str, Any] = {
            "cpf_matches": len(matched_by_cpf),
            "name_matches": len(matched_by_nome),
            "email_matches": len(matched_by_email),
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
        # busca a empresa, centro de custo e descrição que estiver configurado no usuário.
        out_df["CodigoCCustoEmpresa"] = pick(
            matched,
            "Codigo_Centro_de_Custo",
        )
        out_df["DescricaoCCustoEmpresa"] = pick(matched, "Centro_de_Custo")
        out_df["ViajanteMasterNacional"] = pick(matched, "ViajanteMasterNacional")
        out_df["ViajanteMasterInternacional"] = pick(matched, "ViajanteMasterInternacional")
        out_df["EmpresaCCustoParaUsuario"] = "S"
        out_df["Terceiro"] = pick(matched, "Terceiro")
        out_df["CodigoIntegracao"] = "AUT"
        out_df["Status"] = ""

        # Valores padrão para campos que serão preenchidos com mapeamento booleano
        bool_defaults = [
            "Solicitante",
            "Vip",
            "ViajanteMasterNacional",
            "ViajanteMasterInternacional",
            "SolicitanteMaster",
            "MasterAdiantamento",
            "MasterReembolso",
            "Terceiro",
        ]
        for c in ["Endereco", "Cidade", "Estado", "CEP"]:
            out_df[c] = ""
        for c in bool_defaults:
            out_df[c] = "N"

        # Preencher flags booleanas a partir das colunas reais na base (detecção robusta).
        # Igualdade normalizada (sem espaços/_/-): evita que 'Solicitante' case com
        # uma coluna 'SolicitanteMaster' por conter a substring.
        def _norm_col(s):
            return upper_no_accents(str(s)).replace(" ", "").replace("_", "").replace("-", "").strip()

        def find_real_col(target_name):
            t = _norm_col(target_name)
            for k, v in col_map.items():
                if _norm_col(k) == t:
                    return v
            return None

        bool_map = {
            "S": "S",
            "SIM": "S",
            "YES": "S",
            "Y": "S",
            "TRUE": "S",
            "1": "S",
            "N": "N",
            "NAO": "N",
            "NÃO": "N",
            "NO": "N",
            "FALSE": "N",
            "0": "N",
        }
        for logical_col in bool_defaults:
            real_col = find_real_col(logical_col)
            if real_col and real_col in matched.columns:
                vals = (
                    matched[real_col]
                    .fillna("")
                    .astype(str)
                    .map(lambda value: bool_map.get(upper_no_accents(value).strip().upper(), "N"))
                )
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
                inactive["cpf"] = matched_by_cpf.fillna("").to_dict(orient="records")
            else:
                inactive["cpf"] = []
        except Exception:
            inactive["cpf"] = []
        try:
            if not matched_by_nome.empty:
                inactive["nome"] = matched_by_nome.fillna("").to_dict(orient="records")
            else:
                inactive["nome"] = []
        except Exception:
            inactive["nome"] = []
        try:
            if not matched_by_email.empty:
                inactive["email"] = matched_by_email.fillna("").to_dict(orient="records")
            else:
                inactive["email"] = []
        except Exception:
            inactive["email"] = []

        stats["inactive_matches"] = inactive
        # total de matches combinados (fonte de verdade para contagem no preview)
        try:
            stats["total_matches"] = int(matched.shape[0])
        except Exception:
            stats["total_matches"] = sum(len(v) for v in inactive.values() if isinstance(v, list))

        logger.info(
            f"Inativação concluída. Linhas encontradas: {out_df.shape[0]} (CPF={stats['cpf_matches']}, Nome={stats['name_matches']}, Email={stats['email_matches']})"
        )

        return out_df, stats

    except Exception as e:
        logger.error(f"Erro em processar_inativacao_from_paths: {e}")
        # Não converter um bug interno em "nenhum resultado encontrado": propaga
        # para que o handler da rota responda 500 (erro real) em vez de 400
        # (mensagem de negócio "nada encontrado"), que enganaria o usuário.
        raise
