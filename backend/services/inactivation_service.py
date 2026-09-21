import re

import pandas as pd

from backend.core.logging import get_logger
from backend.processor import processar_inativacao_from_paths
from backend.shared.cpf_utils import clean_cpf
from backend.shared.text_utils import upper_no_accents

logger = get_logger()


class InactivationService:
    @staticmethod
    def build_lista_from_text(lista_text: str) -> pd.DataFrame:
        """Constrói o DataFrame de 'lista' a partir de texto colado (uma linha
        por CPF/nome/e-mail). Usado por /process_inativacao e /preview_inativacao
        para que os dois fluxos classifiquem cada linha da mesma forma."""
        lista_items = [item.strip() for item in (lista_text or "").split("\n") if item.strip()]
        email_pat = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.IGNORECASE)
        rows = []
        for it in lista_items:
            digits = re.sub(r"\D", "", it)
            if email_pat.match(it):
                rows.append({"CPF": "", "NomeCompleto": "", "Email": it})
            elif len(digits) == 11:
                rows.append({"CPF": it, "NomeCompleto": "", "Email": ""})
            else:
                rows.append({"CPF": "", "NomeCompleto": it, "Email": ""})
        df_lista = pd.DataFrame(rows)
        if "Email" not in df_lista.columns:
            df_lista["Email"] = ""
        return df_lista

    @staticmethod
    def normalize_lista_columns(df_lista: pd.DataFrame) -> pd.DataFrame:
        try:
            norm_map = {upper_no_accents(re.sub(r"\s+", " ", str(c))).strip(): c for c in df_lista.columns}

            cpf_src = next((v for k, v in norm_map.items() if "CPF" in k), None)
            if "CPF" not in df_lista.columns:
                df_lista["CPF"] = df_lista[cpf_src] if cpf_src else ""

            nome_src = next((v for k, v in norm_map.items() if "NOME COMPLETO" in k or "NOMECOMPLETO" in k), None)
            if "NomeCompleto" not in df_lista.columns:
                if nome_src:
                    df_lista["NomeCompleto"] = df_lista[nome_src]
                else:
                    nome = next((v for k, v in norm_map.items() if k == "NOME" or k.endswith(" NOME")), None)
                    sobrenome = next((v for k, v in norm_map.items() if "SOBRENOME" in k), None)
                    if nome and sobrenome:
                        df_lista["NomeCompleto"] = (
                            df_lista[nome].astype(str).fillna("") + " " + df_lista[sobrenome].astype(str).fillna("")
                        ).str.strip()
                    elif nome:
                        df_lista["NomeCompleto"] = df_lista[nome]
                    else:
                        any_nome = next((v for k, v in norm_map.items() if "NOME" in k), None)
                        df_lista["NomeCompleto"] = df_lista[any_nome] if any_nome else ""

            for col in ["CPF", "NomeCompleto"]:
                if col in df_lista.columns:
                    df_lista[col] = df_lista[col].astype(str).fillna("")

            email_src = next((v for k, v in norm_map.items() if "EMAIL" in k), None)
            if "Email" not in df_lista.columns:
                df_lista["Email"] = df_lista[email_src] if email_src else ""
            if "Email" in df_lista.columns:
                df_lista["Email"] = df_lista["Email"].astype(str).fillna("").str.strip()
        except Exception as exc:
            logger.warning("Normalização de colunas da lista falhou: %s", exc)

        return df_lista

    @staticmethod
    def _detect_base_cols(df_base: pd.DataFrame):
        col_map = {upper_no_accents(str(c)).strip(): c for c in df_base.columns}
        cpf_col = next((v for k, v in col_map.items() if "CPF" in k), None)
        nome_col = next(
            (v for k, v in col_map.items() if "NOMECOMPLETO" in k or "NOME COMPLETO" in k or k == "NOME COMPLETO"), None
        )
        email_col = next((v for k, v in col_map.items() if "EMAIL" in k), None)
        status_col = next((v for k, v in col_map.items() if "STATUS" in k), None)
        userid_col = next(
            (v for k, v in col_map.items() if "USERID" in k or "IDUSUARIO" in k or k.endswith(" USERID")), None
        )
        return cpf_col, nome_col, email_col, status_col, userid_col

    @staticmethod
    def search_matches(df_base: pd.DataFrame, items: list[str]) -> dict:
        raw_items = [str(x).strip() for x in (items or [])]
        cpf_digits = [re.sub(r"\D", "", x) for x in raw_items]
        valid_cpfs = [x for x in cpf_digits if len(x) == 11]

        email_pat = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.IGNORECASE)
        valid_emails = [x for x in raw_items if email_pat.match(x) and re.sub(r"\D", "", x) not in valid_cpfs]

        def is_valid_fullname(name: str) -> bool:
            s = upper_no_accents(str(name)).strip()
            parts = s.split()
            return len(parts) >= 2 and len(s) >= 3

        valid_names_raw = [
            x
            for x in raw_items
            if re.sub(r"\D", "", x) not in valid_cpfs and x not in valid_emails and is_valid_fullname(x)
        ]
        valid_names_norm = [upper_no_accents(x).strip() for x in valid_names_raw]

        seen = set()
        duplicates = []
        for cpf in valid_cpfs:
            if cpf in seen and cpf not in duplicates:
                duplicates.append(cpf)
            seen.add(cpf)

        cpf_col, nome_col, email_col, status_col, userid_col = InactivationService._detect_base_cols(df_base)

        frame = df_base.copy()
        frame["CPFdigits"] = frame[cpf_col].apply(clean_cpf) if cpf_col else ""
        frame["NomeNorm"] = frame[nome_col].apply(lambda v: upper_no_accents(str(v)).strip()) if nome_col else ""

        results = []
        found_cpfs = set()
        found_emails = set()
        if valid_cpfs and cpf_col:
            matches = frame[frame["CPFdigits"].isin(valid_cpfs)].copy()
            for _, row in matches.iterrows():
                cpf = row.get("CPFdigits", "")
                found_cpfs.add(cpf)
                results.append(
                    {
                        "id": str(row.get(userid_col, "")) if userid_col and row.get(userid_col, "") != "" else None,
                        "nome": str(row.get(nome_col, "")) if nome_col else str(row.get("NomeCompleto", "")),
                        "cpf": cpf,
                        "email": str(row.get(email_col, "")) if email_col else str(row.get("Email", "")),
                        "status_atual": str(row.get(status_col, "")) if status_col else str(row.get("Status", "")),
                        "found": True,
                    }
                )

        found_name_norms = set()
        if valid_names_norm and nome_col:
            matches = frame[frame["NomeNorm"].isin(valid_names_norm)].copy()
            for _, row in matches.iterrows():
                cpf = row.get("CPFdigits", "")
                name_norm = row.get("NomeNorm", "")
                if cpf and cpf in found_cpfs:
                    found_name_norms.add(name_norm)
                    continue
                found_name_norms.add(name_norm)
                results.append(
                    {
                        "id": str(row.get(userid_col, "")) if userid_col and row.get(userid_col, "") != "" else None,
                        "nome": str(row.get(nome_col, "")) if nome_col else str(row.get("NomeCompleto", "")),
                        "cpf": cpf,
                        "email": str(row.get(email_col, "")) if email_col else str(row.get("Email", "")),
                        "status_atual": str(row.get(status_col, "")) if status_col else str(row.get("Status", "")),
                        "found": True,
                    }
                )

        if valid_emails and email_col:
            base_email_norm = frame[email_col].astype(str).fillna("").str.strip().str.lower()
            target_emails_norm = [email.strip().lower() for email in valid_emails]
            matches = frame[base_email_norm.isin(target_emails_norm)].copy()
            for _, row in matches.iterrows():
                email_val = str(row.get(email_col, "")).strip()
                found_emails.add(email_val.lower())
                results.append(
                    {
                        "id": str(row.get(userid_col, "")) if userid_col and row.get(userid_col, "") != "" else None,
                        "nome": str(row.get(nome_col, "")) if nome_col else str(row.get("NomeCompleto", "")),
                        "cpf": str(row.get("CPFdigits", "")),
                        "email": email_val,
                        "status_atual": str(row.get(status_col, "")) if status_col else str(row.get("Status", "")),
                        "found": True,
                    }
                )

        not_found_cpfs = [cpf for cpf in valid_cpfs if cpf not in found_cpfs]
        for cpf in not_found_cpfs:
            results.append(
                {"id": None, "nome": "", "cpf": cpf, "email": "", "status_atual": "Não localizado", "found": False}
            )

        not_found_names_raw = [
            raw for raw, norm in zip(valid_names_raw, valid_names_norm) if norm not in found_name_norms
        ]
        for raw, norm in zip(valid_names_raw, valid_names_norm):
            if norm not in found_name_norms:
                results.append(
                    {"id": None, "nome": raw, "cpf": "", "email": "", "status_atual": "Não localizado", "found": False}
                )

        not_found_emails = [email for email in valid_emails if email.strip().lower() not in found_emails]
        for email in not_found_emails:
            results.append(
                {"id": None, "nome": "", "cpf": "", "email": email, "status_atual": "Não localizado", "found": False}
            )

        results = sorted(
            results, key=lambda item: (0 if item.get("found") else 1, item.get("nome") or "", item.get("cpf") or "")
        )
        return {
            "items": results,
            "total": len(results),
            "duplicates": duplicates,
            "not_found": not_found_cpfs + not_found_names_raw + not_found_emails,
        }

    @staticmethod
    def process_from_dataframes(df_base: pd.DataFrame, df_lista: pd.DataFrame):
        return processar_inativacao_from_paths(df_base, df_lista)
