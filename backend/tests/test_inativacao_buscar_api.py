"""Testes de /api/inativacao/buscar (sem cobertura de API antes desta suíte;
InactivationService.search_matches só era exercitado unitariamente)."""

import pandas as pd
from _helpers import valid_cpf, xlsx_upload


def _base_df():
    return pd.DataFrame(
        [
            {"CPF": valid_cpf(1), "NomeCompleto": "Ana Souza", "Email": "ana@x.com", "Status": "ATIVO"},
            {"CPF": valid_cpf(2), "NomeCompleto": "Bruno Lima", "Email": "bruno@x.com", "Status": "ATIVO"},
        ]
    )


def test_buscar_via_json_itens(client):
    data = {"base": xlsx_upload(_base_df(), "base.xlsx")}
    resp = client.post(
        "/api/inativacao/buscar",
        data={**data, "itens": f'["{valid_cpf(1)}"]'},
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["total"] == 1
    assert body["items"][0]["found"] is True
    assert body["items"][0]["nome"] == "Ana Souza"


def test_buscar_via_lista_text(client):
    data = {
        "base": xlsx_upload(_base_df(), "base.xlsx"),
        "lista_text": f"{valid_cpf(1)}\n{valid_cpf(2)}",
    }
    resp = client.post("/api/inativacao/buscar", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["total"] == 2
    assert all(item["found"] for item in body["items"])


def test_buscar_via_lista_file_extracts_cpf_column(client):
    data = {
        "base": xlsx_upload(_base_df(), "base.xlsx"),
        "lista": xlsx_upload(pd.DataFrame([{"CPF": valid_cpf(1)}]), "lista.xlsx"),
    }
    resp = client.post("/api/inativacao/buscar", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["total"] == 1
    assert body["items"][0]["found"] is True


def test_buscar_not_found_item(client):
    data = {"base": xlsx_upload(_base_df(), "base.xlsx"), "lista_text": valid_cpf(3)}
    resp = client.post("/api/inativacao/buscar", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["items"][0]["found"] is False
    assert body["items"][0]["status_atual"] == "Não localizado"
    assert body["not_found"] == [valid_cpf(3)]


def test_buscar_reports_duplicate_cpf_in_input_list(client):
    cpf = valid_cpf(1)
    data = {"base": xlsx_upload(_base_df(), "base.xlsx"), "lista_text": f"{cpf}\n{cpf}"}
    resp = client.post("/api/inativacao/buscar", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert resp.get_json()["duplicates"] == [cpf]


def test_buscar_without_base_file_rejected(client):
    resp = client.post(
        "/api/inativacao/buscar", data={"lista_text": valid_cpf(1)}, content_type="multipart/form-data"
    )
    assert resp.status_code == 400
