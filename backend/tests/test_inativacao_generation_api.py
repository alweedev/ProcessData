"""Geração de inativação: função e endpoint.

Verificações portadas de backend/test_inativacao_run.py e
backend/test_integration_api.py (scripts manuais removidos em D5).
"""

import pandas as pd
from _helpers import valid_cpf, xlsx_upload

from backend.processor import processar_inativacao_from_paths


def _base():
    return pd.DataFrame(
        [
            {"CPF": valid_cpf(1), "NomeCompleto": "Maria Silva", "Email": "maria@x.com", "Status": "ATIVO"},
            {"CPF": valid_cpf(2), "NomeCompleto": "Joao Pereira", "Email": "joao@x.com", "Status": "ATIVO"},
            {"CPF": valid_cpf(3), "NomeCompleto": "Mariana Costa", "Email": "mc@x.com", "Status": "INATIVO"},
        ]
    )


def test_match_by_exact_cpf():
    out, stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"CPF": valid_cpf(1)}]))
    assert stats["cpf_matches"] == 1
    assert out["NomeCompleto"].tolist() == ["MARIA SILVA"]
    assert out["Operacao"].tolist() == ["DELETE"]


def test_match_by_exact_full_name():
    out, stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"NomeCompleto": "Maria Silva"}]))
    assert stats["name_matches"] == 1
    assert out["NomeCompleto"].tolist() == ["MARIA SILVA"]


def test_generic_single_name_does_not_over_match():
    out, _stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"NomeCompleto": "Maria"}]))
    # "Maria" != "Maria Silva" nem "Mariana Costa" -> nenhum match exato
    assert out.empty


def test_inactive_row_is_not_matched():
    out, stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"CPF": valid_cpf(3)}]))
    assert out.empty
    assert stats["cpf_matches"] == 0


def test_process_inativacao_endpoint_returns_xlsx(client):
    data = {
        "base": xlsx_upload(_base(), "base.xlsx"),
        "lista": xlsx_upload(pd.DataFrame([{"CPF": valid_cpf(1)}]), "lista.xlsx"),
    }
    resp = client.post("/api/process_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert "spreadsheetml" in resp.headers.get("Content-Type", "")
    assert resp.data[:2] == b"PK"


def test_process_inativacao_endpoint_400_when_no_match(client):
    data = {
        "base": xlsx_upload(_base(), "base.xlsx"),
        "lista": xlsx_upload(pd.DataFrame([{"CPF": valid_cpf(3)}]), "lista.xlsx"),  # INATIVO
    }
    resp = client.post("/api/process_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400


def test_internal_bug_is_not_reported_as_business_no_match(client, monkeypatch):
    """Um bug interno em processar_inativacao_from_paths deve virar 500
    (erro real), nunca o 400 de 'nenhum dado encontrado' (que enganaria o
    usuário fazendo parecer que a base/lista simplesmente não bateram)."""
    from backend.services.inactivation_service import InactivationService

    def boom(*a, **k):
        raise RuntimeError("bug interno inesperado")

    monkeypatch.setattr(InactivationService, "process_from_dataframes", boom)

    data = {
        "base": xlsx_upload(_base(), "base.xlsx"),
        "lista": xlsx_upload(pd.DataFrame([{"CPF": valid_cpf(1)}]), "lista.xlsx"),
    }
    resp = client.post("/api/process_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 500
    assert "interno" in resp.get_json()["error"].lower()
