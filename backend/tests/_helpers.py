"""Utilitários de teste: geração de planilhas em memória e CPFs válidos."""

import io

import pandas as pd


def xlsx_bytes(df: pd.DataFrame) -> io.BytesIO:
    """Serializa um DataFrame como .xlsx em memória (ponteiro no início)."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


def xlsx_upload(df: pd.DataFrame, name: str = "planilha.xlsx"):
    """Tupla ``(BytesIO, filename)`` pronta para ``client.post(..., data=...)``."""
    return (xlsx_bytes(df), name)


def _cpf_check_digits(base9: str) -> str:
    digits = [int(c) for c in base9]
    dv1_sum = sum(d * w for d, w in zip(digits, range(10, 1, -1)))
    r = dv1_sum % 11
    dv1 = 0 if r < 2 else 11 - r
    digits.append(dv1)
    dv2_sum = sum(d * w for d, w in zip(digits, range(11, 1, -1)))
    r = dv2_sum % 11
    dv2 = 0 if r < 2 else 11 - r
    return f"{dv1}{dv2}"


def valid_cpf(seed: int = 0) -> str:
    """CPF de 11 dígitos com dígito verificador válido (determinístico por seed).

    A base de 9 dígitos fica entre 100000000 e 899999999 para evitar zeros à
    esquerda (que planilhas costumam coagir para número).
    """
    n = 100_000_000 + (seed * 7_654_321) % 800_000_000
    base = f"{n:09d}"
    if len(set(base)) == 1:  # evita CPFs de dígitos repetidos (inválidos por regra)
        base = "123456789"
    return base + _cpf_check_digits(base)


def format_cpf(cpf_digits: str) -> str:
    """``12345678909`` -> ``123.456.789-09``."""
    d = cpf_digits
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
