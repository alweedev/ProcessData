from backend.domain.rules import REQUIRED_OUTPUT_COLS
from backend.shared.cpf_utils import clean_cpf
from backend.shared.text_utils import upper_no_accents


class ValidationService:
    # Colunas obrigatórias em que uma coluna inteira vazia indica que a origem
    # não trouxe o dado (sinal útil para o relatório). CC código/descrição
    # ficam de fora — podem ser legitimamente vazios em algumas fichas.
    _NON_EMPTY_REQUIRED = [
        "Login", "Email", "NomeCompleto", "Nome", "SobreNome", "NomeEmpresa",
    ]

    @staticmethod
    def validate_dataframe(df) -> list:
        """Valida colunas obrigatórias de saída (erro __geral__)."""
        msgs = []
        for col in REQUIRED_OUTPUT_COLS:
            if col not in df.columns:
                msgs.append(f"Coluna obrigatoria ausente: {col}")
        if df.shape[0] > 0:
            for col in ValidationService._NON_EMPTY_REQUIRED:
                if col in df.columns and df[col].astype(str).str.strip().eq("").all():
                    msgs.append(f"Coluna obrigatoria vazia: {col}")
        return msgs

    @staticmethod
    def validate_row(row):
        errors = []

        solicitante = str(row.get("Solicitante", "")).strip().upper()
        if solicitante not in ("S", "N"):
            errors.append("Solicitante obrigatório (deve ser S ou N)")

        cpf_raw = row.get("CPF", "") or row.get("Login", "")
        cpf_digits = clean_cpf(cpf_raw)
        if cpf_raw and cpf_digits and len(cpf_digits) != 11:
            errors.append("CPF deve ter 11 dígitos")

        email = str(row.get("Email", "")).strip()
        if email and ("@" not in email or "." not in email.split("@")[-1]):
            errors.append("Email inválido")

        if not str(row.get("NomeCompleto", "")).strip():
            errors.append("NomeCompleto vazio")

        nivel = upper_no_accents(row.get("Nivel", ""))
        if nivel and nivel not in ("OPERACIONAL", "GERENCIA", "DIRETORIA"):
            if "OPER" in nivel:
                row["Nivel"] = "OPERACIONAL"
            elif "GER" in nivel:
                row["Nivel"] = "GERENCIA"
            elif "DIR" in nivel:
                row["Nivel"] = "DIRETORIA"
            else:
                row["Nivel"] = ""
                errors.append("Nivel inválido, ajustado para vazio")

        return errors
