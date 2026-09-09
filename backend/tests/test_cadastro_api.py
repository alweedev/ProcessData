"""Testes do endpoint de cadastro (/api/process_cadastro)."""
import io

import pandas as pd

from _helpers import valid_cpf, xlsx_upload


def test_docx_upload_rejected(client):
    data = {"files[]": (io.BytesIO(b"PK\x03\x04 fake docx"), "ficha.docx")}
    resp = client.post("/api/process_cadastro", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert "extens" in resp.get_json()["error"].lower()


def test_cadastro_happy_path_returns_xlsx(client):
    df = pd.DataFrame(
        [
            {"CPF": valid_cpf(10), "NOME COMPLETO": "Ana Souza", "EMAIL": "ana@empresa.com",
             "EMPRESA": "Empresa A", "Centro de custo": "CC1", "SOLICITANTE? (S/N)": "S"},
        ]
    )
    data = {"files[]": xlsx_upload(df, "cadastro.xlsx"), "login_choice": "CPF", "fluxo": "SELF"}
    resp = client.post("/api/process_cadastro", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert "spreadsheetml" in resp.headers.get("Content-Type", "")
    assert resp.data[:2] == b"PK"  # xlsx é um zip
