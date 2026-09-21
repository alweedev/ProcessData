from backend.shared.cpf_utils import clean_cpf


def mascarar_cpf(value) -> str:
    """``12345678909`` -> ``***.456.789-**`` (só os 6 dígitos do meio); ``""`` se não houver CPF de 11 dígitos."""
    digits = clean_cpf(value)
    if len(digits) != 11:
        return ""
    return f"***.{digits[3:6]}.{digits[6:9]}-**"
