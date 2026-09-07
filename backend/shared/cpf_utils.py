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
