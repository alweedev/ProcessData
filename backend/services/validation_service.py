import re

from backend.domain.rules import REQUIRED_FICHA_FIELDS
from backend.shared.cpf_utils import clean_cpf, raw_cpf_digits
from backend.shared.text_utils import upper_no_accents

# Um endereço só: sem espaço, ";" ou ","; um "@"; parte local não vazia; domínio com ponto,
# sem rótulo vazio ("a@b..com", "a@.com") e TLD de 2+ caracteres ("a@b.", "a@b.c" não passam).
_EMAIL_RE = re.compile(r"^[^\s@;,]+@(?:[^\s@;,.]+\.)+[^\s@;,.]{2,}$")


def _blank(value) -> bool:
    return not str(value if value is not None else "").strip()


def cpf_is_waived(row, login_choice: str = "CPF") -> bool:
    """Estrangeiro sem CPF: com número de passaporte na linha, o CPF deixa de ser obrigatório —
    exceto no login por CPF, em que o Login é gerado a partir dele."""
    return login_choice != "CPF" and not _blank(row.get("Passaporte", ""))


class ValidationService:
    @staticmethod
    def validate_dataframe(df, ausentes=(), login_choice: str = "CPF") -> list:
        """Erros gerais da ficha (``__geral__``) sobre os campos obrigatórios.

        ``ausentes``: campos obrigatórios que a ficha nem trazia como coluna (o pipeline
        completa as colunas faltantes com vazio antes de validar, então só ele sabe).
        Uma coluna obrigatória presente mas vazia em TODAS as linhas indica que a origem
        não trouxe o dado — sinal mais útil que uma lista de linhas iguais.
        """
        msgs = []
        for campo in ausentes:
            msgs.append(f"Coluna obrigatória ausente na ficha: {REQUIRED_FICHA_FIELDS[campo]}")
        if df.shape[0] > 0:
            # CPF vazio na ficha inteira é esperado quando todos são estrangeiros com passaporte.
            cpf_dispensavel = (
                login_choice != "CPF"
                and "Passaporte" in df.columns
                and df["Passaporte"].astype(str).str.strip().ne("").any()
            )
            for campo, rotulo in REQUIRED_FICHA_FIELDS.items():
                if campo in ausentes or campo not in df.columns:
                    continue
                if campo == "CPF" and cpf_dispensavel:
                    continue
                if df[campo].astype(str).str.strip().eq("").all():
                    msgs.append(f"Coluna obrigatória vazia na ficha: {rotulo}")
        return msgs

    @staticmethod
    def validate_row(row, login_choice: str = "CPF"):
        errors = []

        # Obrigatórios: em branco = linha inválida. Solicitante/Terceiro e os demais campos de
        # sim/não NÃO entram aqui: o pipeline os normaliza depois (Sim -> S, N -> N, branco -> N).
        waive_cpf = cpf_is_waived(row, login_choice)
        for campo, rotulo in REQUIRED_FICHA_FIELDS.items():
            if campo == "CPF" and waive_cpf:
                continue
            if _blank(row.get(campo, "")):
                errors.append(f"Campo obrigatório em branco: {rotulo}")

        cpf_raw = row.get("CPF", "")
        cpf_digits = clean_cpf(cpf_raw)
        if cpf_raw and cpf_digits:
            # zfill restaura no máximo 1 zero à esquerda perdido pelo Excel
            # (10 -> 11 dígitos); menos que isso é entrada incompleta, não
            # CPF válido com zero suprimido.
            if len(cpf_digits) != 11 or len(raw_cpf_digits(cpf_raw)) < 10:
                errors.append("CPF deve ter 11 dígitos")

        email = str(row.get("Email", "")).strip()
        if email and not _EMAIL_RE.match(email):
            errors.append("Email inválido")

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
