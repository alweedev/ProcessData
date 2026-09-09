"""Testes dos endpoints de inativação: preview e rota /executar removida (B1)."""
import os

import pandas as pd

from backend.core.config import settings

from _helpers import valid_cpf, xlsx_upload


def _base_df():
    return pd.DataFrame(
        [
            {"CPF": valid_cpf(1), "NomeCompleto": "Ana Souza", "Email": "ana@x.com", "Status": "ATIVO"},
            {"CPF": valid_cpf(2), "NomeCompleto": "Bruno Lima", "Email": "bruno@x.com", "Status": "ATIVO"},
        ]
    )


def test_preview_inativacao_ok(client):
    base = _base_df()
    lista = pd.DataFrame([{"CPF": valid_cpf(1)}])
    data = {
        "base": xlsx_upload(base, "base.xlsx"),
        "lista": xlsx_upload(lista, "lista.xlsx"),
    }
    resp = client.post("/api/preview_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["count"] >= 1
    assert body["columns"]
    assert body["records"]
    assert isinstance(body["stats"], dict)
    # temporários limpos
    assert os.listdir(settings.UPLOAD_FOLDER) == []


def test_preview_inativacao_bad_extension(client):
    data = {
        "base": (xlsx_upload(_base_df())[0], "base.txt"),
        "lista_text": valid_cpf(1),
    }
    resp = client.post("/api/preview_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400


def test_inativacao_executar_gone(client):
    resp = client.post("/api/inativacao/executar", json={"usuarios": []})
    # rota removida: 404 (sem regra) ou 405 (só o catch-all de SPA aceita GET)
    assert resp.status_code in (404, 405)
    assert resp.get_json() != {"success": True}
