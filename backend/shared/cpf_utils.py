import re


def clean_cpf(value):
    if value is None:
        return ""
    return re.sub(r"\D", "", str(value))


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
