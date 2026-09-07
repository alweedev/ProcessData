import pandas as pd


class ReportService:
    @staticmethod
    def build_quality_report(df: pd.DataFrame, errors: dict) -> dict:
        errors = errors or {}
        line_errors = {k: v for k, v in errors.items() if isinstance(k, int)}
        general_errors = errors.get("__geral__")

        total_rows = int(df.shape[0]) if df is not None else 0
        invalid_rows = len(line_errors)
        valid_rows = max(0, total_rows - invalid_rows)

        duplicated_rows = 0
        if df is not None and not df.empty and "Login" in df.columns and "NomeCompleto" in df.columns:
            duplicated_rows = int(df.duplicated(subset=["Login", "NomeCompleto"]).sum())

        required_blank = {}
        if df is not None and not df.empty:
            required = [
                "Login",
                "NomeEmpresa",
                "CodigoCCustoEmpresa",
                "DescricaoCCustoEmpresa",
                "Email",
                "NomeCompleto",
            ]
            for col in required:
                if col in df.columns:
                    required_blank[col] = int(df[col].astype(str).str.strip().eq("").sum())

        return {
            "total_rows": total_rows,
            "valid_rows": valid_rows,
            "invalid_rows": invalid_rows,
            "duplicated_rows": duplicated_rows,
            "general_errors": general_errors or "",
            "line_errors": line_errors,
            "required_blank": required_blank,
        }
