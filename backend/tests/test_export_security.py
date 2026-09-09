"""Neutralização de formula injection no Excel exportado (P8a)."""

import io

import pandas as pd
from _helpers import valid_cpf, xlsx_upload
from openpyxl import load_workbook

from backend.services.export_service import ExportService

PAYLOADS = ["=SUM(A1)", "+1", "-1", "@x", "\tx"]
_STRING_TYPES = {"s", "str", "inlineStr"}  # qualquer um != fórmula


def test_formula_cells_written_as_text():
    out = ExportService.to_excel_bytes(pd.DataFrame({"x": PAYLOADS}), sheet_name="S")
    ws = load_workbook(io.BytesIO(out.getvalue()))["S"]
    for i, expected in enumerate(PAYLOADS, start=2):  # linha 1 = cabeçalho
        cell = ws.cell(row=i, column=1)
        assert cell.data_type in _STRING_TYPES, f"{expected!r} veio como {cell.data_type}"
        assert cell.data_type != "f"
        assert cell.value == expected


def test_normal_values_unchanged():
    out = ExportService.to_excel_bytes(pd.DataFrame({"x": ["ABC", "Joao", "123"]}), sheet_name="S")
    ws = load_workbook(io.BytesIO(out.getvalue()))["S"]
    assert ws.cell(row=2, column=1).value == "ABC"


def test_aprovacao_export_neutralizes_injection(client):
    users = pd.DataFrame([{"CPF": valid_cpf(1), "NomeCompleto": "Aprovador Um", "Status": "ATIVO"}])
    # payload com prefixo "+": sobrevive ao round-trip de leitura da planilha
    # como string (um "=" seria lido como fórmula sem valor em cache).
    base = pd.DataFrame(
        [
            {
                "AprovacaoId": "A",
                "AprovacaoPor": "+cmd|' /c calc'!A0",
                "LoginAprovador_1": valid_cpf(1),
                "LoginAprovador_2": valid_cpf(3),
            }
        ]
    )
    data = {
        "users_file": xlsx_upload(users, "u.xlsx"),
        "base_file": xlsx_upload(base, "b.xlsx"),
        "cpf": valid_cpf(1),
        "mode": "all",
        "ignore_empty_warning": "true",
    }
    resp = client.post("/api/aprovacao/remover/export", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    ws = load_workbook(io.BytesIO(resp.data)).active
    header = [c.value for c in ws[1]]
    col = header.index("AprovacaoPor") + 1
    cell = ws.cell(row=2, column=col)
    assert cell.data_type in _STRING_TYPES  # nunca "f" (fórmula)
    assert str(cell.value).startswith("+cmd")
