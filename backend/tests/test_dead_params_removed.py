"""Remoção dos parâmetros mortos use_fuzzy/fuzzy_cutoff (C3)."""

import inspect

import pandas as pd
from _helpers import valid_cpf, xlsx_upload

from backend.processor import processar_inativacao_from_paths
from backend.services.inactivation_service import InactivationService


def test_signatures_have_no_fuzzy_params():
    for fn in (processar_inativacao_from_paths, InactivationService.process_from_dataframes):
        params = set(inspect.signature(fn).parameters)
        assert "use_fuzzy" not in params
        assert "fuzzy_cutoff" not in params
    assert set(inspect.signature(InactivationService.process_from_dataframes).parameters) == {
        "df_base",
        "df_lista",
    }


def test_process_inativacao_still_works(client):
    base = pd.DataFrame([{"CPF": valid_cpf(1), "NomeCompleto": "Ana Souza", "Email": "a@x.com", "Status": "ATIVO"}])
    lista = pd.DataFrame([{"CPF": valid_cpf(1)}])
    data = {"base": xlsx_upload(base, "base.xlsx"), "lista": xlsx_upload(lista, "lista.xlsx")}
    resp = client.post("/api/process_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert resp.data[:2] == b"PK"
