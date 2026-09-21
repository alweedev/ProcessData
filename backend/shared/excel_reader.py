"""Leitura da 1ª aba de uma planilha como texto (tudo `str`, como o pipeline de cadastro espera)."""

import pandas as pd


def _read_first_sheet(path: str, engine: str | None) -> tuple[pd.DataFrame, str, int]:
    with pd.ExcelFile(path, engine=engine) as workbook:
        if not workbook.sheet_names:
            raise ValueError("A planilha não tem nenhuma aba.")
        first = str(workbook.sheet_names[0])
        return workbook.parse(sheet_name=first, dtype=str), first, len(workbook.sheet_names)


def read_first_sheet_as_text(path: str) -> tuple[pd.DataFrame, str, int]:
    """Devolve ``(dados, nome da 1ª aba, nº de abas)``.

    `.xlsx` é lido pelo calamine, ~6x mais rápido que o openpyxl e com células idênticas (conferido
    nos modelos de ficha dos clientes e em todos os tipos de célula: texto, número, data, vazio...).
    Qualquer falha cai no motor padrão do pandas, cujas mensagens de erro são as que o projeto sempre
    mostrou. `.xls` segue no xlrd.
    """
    if path.lower().endswith(".xlsx"):
        try:
            return _read_first_sheet(path, "calamine")
        except Exception:
            pass
    return _read_first_sheet(path, None)
