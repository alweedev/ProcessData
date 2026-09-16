import re


def raw_cpf_digits(value):
    """Dígitos de ``value`` antes do zero-padding de `clean_cpf`.

    Usado para diferenciar "Excel perdeu 1 zero à esquerda" (10 dígitos reais)
    de entrada claramente incompleta (poucos dígitos) — ambas viram 11 dígitos
    depois do `zfill`, mas só a primeira é uma restauração legítima.
    """
    if value is None:
        return ""
    text = str(value).strip()
    # openpyxl/pandas às vezes devolvem célula numérica como float ("...909.0");
    # sem isso o "0" espúrio vira um 12º dígito.
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return re.sub(r"\D", "", text)


def clean_cpf(value):
    digits = raw_cpf_digits(value)
    # Planilhas tratando CPF como número perdem zeros à esquerda; restaura o
    # tamanho padrão de 11 dígitos (mesma defesa usada em processor.py).
    return digits.zfill(11) if digits else ""


# Alias histórico (backend.utils.limpar_cpf_raw)
limpar_cpf_raw = clean_cpf


def format_cpf_for_output(cpf_digits):
    if not cpf_digits:
        return ""
    digits = re.sub(r"\D", "", str(cpf_digits))
    if len(digits) == 11:
        return f"{digits[:-2]}-{digits[-2:]}"
    return digits


def is_valid_cpf(value) -> bool:
    """Valida CPF pelos dígitos verificadores (módulo 11).

    Retorna ``False`` para entradas sem 11 dígitos, com todos os dígitos iguais
    (ex.: ``00000000000``) ou com dígito verificador incorreto.
    Usado hoje apenas no fluxo de aprovação; cadastro/inativação seguem com
    checagem de comprimento.
    """
    digits = clean_cpf(value)
    if len(digits) != 11 or len(set(digits)) == 1:
        return False

    nums = [int(c) for c in digits]
    for pos in (9, 10):
        weights = range(pos + 1, 1, -1)
        total = sum(n * w for n, w in zip(nums[:pos], weights))
        check = (total * 10) % 11
        if check == 10:
            check = 0
        if check != nums[pos]:
            return False
    return True
