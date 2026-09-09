import os

import pandas as pd

from backend.domain.rules import MODEL_COLS, FICHA_MAP
from backend.services.validation_service import ValidationService
from backend.shared.text_utils import sanitize_output_text, split_name_first_last, upper_no_accents
from backend.shared.cpf_utils import clean_cpf, format_cpf_for_output


class ProcessingService:
    @staticmethod
    def _drop_header_like_rows(df):
        if df.empty:
            return df
        cols = list(df.columns)

        def is_header(row):
            matches = 0
            for col in cols:
                value = str(row.get(col, "")).strip()
                if value.upper() == str(col).upper():
                    matches += 1
            return (matches / max(1, len(cols))) > 0.4

        return df[~df.apply(is_header, axis=1)]

    @staticmethod
    def process_records_from_files(paths, login_choice="CPF", fluxo="SELF"):
        all_data = []
        all_errors = {}

        for path in paths:
            try:
                if path.lower().endswith(".xlsx") or path.lower().endswith(".xls"):
                    df = pd.read_excel(path, dtype=str).fillna("")
                    df = ProcessingService._drop_header_like_rows(df)
                    normalized_map = {upper_no_accents(str(k)).strip(): v for k, v in FICHA_MAP.items()}
                    for _, row in df.iterrows():
                        mapped_row = {}
                        for col in df.columns:
                            normalized_col = upper_no_accents(str(col)).strip()
                            if normalized_col in normalized_map:
                                mapped_row[normalized_map[normalized_col]] = row[col]
                        if mapped_row:
                            all_data.append(mapped_row)
            except Exception as exc:
                all_errors[path] = str(exc)

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
            first, last = split_name_first_last(row.get("NomeCompleto", ""))
            if first:
                df_final.at[idx, "Nome"] = first
            if last:
                df_final.at[idx, "SobreNome"] = last

        if login_choice == "CPF":
            if "CPF" in df_final.columns:
                df_final["Login"] = df_final["CPF"].apply(lambda x: format_cpf_for_output(clean_cpf(x)) if x else "")
        elif login_choice == "EMAIL":
            if "Email" in df_final.columns:
                df_final["Login"] = df_final["Email"]

        fluxo_up = (fluxo or "").upper()
        if fluxo_up == "SELF":
            for col in ["Vip", "ViajanteMasterNacional", "ViajanteMasterInternacional", "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]:
                df_final[col] = "N"
        elif fluxo_up == "FRONT":
            df_final["ViajanteMasterNacional"] = "S"
            df_final["ViajanteMasterInternacional"] = "S"
            for col in ["Vip", "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]:
                df_final[col] = "N"
            if "Login" in df_final.columns:
                def _prefix_front(v):
                    if pd.isna(v) or str(v).strip() == "":
                        return v
                    return "FRONT" + str(v).replace(" ", "")
                df_final["Login"] = df_final["Login"].apply(_prefix_front)

        for col in ["Nome", "SobreNome", "NomeCompleto", "NomeEmpresa", "DescricaoCCustoEmpresa", "DescricaoCCustoCliente", "Cargo", "Departamento", "Cidade", "Estado", "Endereco"]:
            if col in df_final.columns:
                if col in ["Nome", "SobreNome"]:
                    df_final[col] = df_final[col].apply(lambda v: sanitize_output_text(v, 20))
                elif col == "NomeCompleto":
                    df_final[col] = df_final[col].apply(lambda v: sanitize_output_text(v, None))
                else:
                    df_final[col] = df_final[col].apply(lambda v: sanitize_output_text(v, None))

        errors = {}
        for idx, row in df_final.iterrows():
            msgs = ValidationService.validate_row(row)
            if msgs:
                errors[idx] = "; ".join(msgs)

        if "Login" in df_final.columns and "NomeCompleto" in df_final.columns:
            df_final = df_final.drop_duplicates(subset=["Login", "NomeCompleto"], keep="first")

        bool_cols = ["Solicitante", "Terceiro", "Vip", "ViajanteMasterNacional", "ViajanteMasterInternacional", "SolicitanteMaster", "MasterAdiantamento", "MasterReembolso"]
        _true_set = {"S", "SIM", "YES", "Y", "TRUE", "1"}

        def _map_bool_sn(value):
            # fold de acento antes de comparar: "Não"/"Sìm" normalizam
            return "S" if upper_no_accents(value).strip().upper() in _true_set else "N"

        def _map_terceiro(value):
            # regra da ficha: se tiver dígitos, mantém os dígitos (id de terceiro);
            # senão mapeia Sim/Não -> S/N
            digits = "".join(filter(str.isdigit, str(value)))
            return digits if digits else _map_bool_sn(value)

        for bc in bool_cols:
            if bc not in df_final.columns:
                df_final[bc] = "N"
            elif bc == "Terceiro":
                df_final[bc] = df_final[bc].fillna("").apply(_map_terceiro)
            else:
                df_final[bc] = df_final[bc].fillna("").apply(_map_bool_sn)

        if "NroMatricula" in df_final.columns:
            df_final["NroMatricula"] = df_final["NroMatricula"].fillna("").apply(lambda v: "".join(filter(str.isdigit, str(v))))

        for col in df_final.columns:
            if df_final[col].dtype == object:
                if col in ("Email", "Telefone"):
                    df_final[col] = df_final[col].fillna("").astype(str).apply(lambda v: v.strip().upper())
                elif col == "Login" and login_choice == "EMAIL":
                    df_final[col] = df_final[col].fillna("").astype(str).apply(lambda v: v.strip().upper())
                else:
                    df_final[col] = df_final[col].fillna("").astype(str).apply(lambda v: sanitize_output_text(v, None))

        def _drop_blank_rows(frame):
            critical = [c for c in ["Login", "NomeCompleto", "CPF", "Email"] if c in frame.columns]
            if not critical:
                return frame
            trimmed = frame[critical].apply(lambda s: s.astype(str).str.strip())
            mask_blank = trimmed.eq("").all(axis=1)
            return frame.loc[~mask_blank].copy()

        df_final = _drop_blank_rows(df_final)
        df_final = df_final[MODEL_COLS]
        return errors, df_final
