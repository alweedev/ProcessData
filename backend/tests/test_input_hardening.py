"""Hardening de entrada: conteúdo de upload + path traversal no frontend (P8c, P8d)."""

import io

import pandas as pd
from _helpers import valid_cpf, xlsx_upload


def test_fake_xlsx_rejected(client):
    data = {"files[]": (io.BytesIO(b"isto nao e um zip / xlsx"), "malicioso.xlsx")}
    resp = client.post("/api/process_cadastro", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert "xlsx" in resp.get_json()["error"].lower()


def test_real_xlsx_accepted(client):
    df = pd.DataFrame(
        [
            {
                "CPF": valid_cpf(1),
                "NOME COMPLETO": "Ana Souza",
                "EMAIL": "a@x.com",
                "EMPRESA": "E",
                "Centro de custo": "C",
                "SOLICITANTE? (S/N)": "S",
            }
        ]
    )
    data = {"files[]": xlsx_upload(df, "c.xlsx"), "login_choice": "CPF", "fluxo": "SELF"}
    resp = client.post("/api/process_cadastro", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)


def test_frontend_path_traversal_blocked(client):
    alvos = [
        "/..%2f..%2fbackend%2fapp.py",
        "/%2e%2e/backend/app.py",
        "/static/../../backend/app.py",
        "/../backend/core/config.py",
    ]
    for url in alvos:
        resp = client.get(url)
        # nunca pode devolver 200 com o código-fonte Python
        assert not (resp.status_code == 200 and b"create_app" in resp.data), url
        assert not (resp.status_code == 200 and b"_DEFAULT_UPLOAD_FOLDER" in resp.data), url


def test_frontend_serves_index(client):
    resp = client.get("/")
    assert resp.status_code == 200
    assert b"ProcessData" in resp.data
