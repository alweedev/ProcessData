"""Dir de upload fora do repo + limpeza de temporários (P7)."""

import os

import pandas as pd
from _helpers import valid_cpf, xlsx_upload

from backend.core.config import _DEFAULT_UPLOAD_FOLDER, settings


def test_default_upload_dir_is_outside_repo():
    # o valor padrão (antes de qualquer override por env/teste) não fica no repo
    assert not os.path.abspath(_DEFAULT_UPLOAD_FOLDER).startswith(os.path.abspath(settings.PROJECT_ROOT) + os.sep)


def test_cadastro_cleans_temp_files(client):
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
    assert os.listdir(settings.UPLOAD_FOLDER) == []


def test_process_inativacao_cleans_temp_files(client):
    base = pd.DataFrame([{"CPF": valid_cpf(1), "NomeCompleto": "Ana Souza", "Email": "a@x.com", "Status": "ATIVO"}])
    lista = pd.DataFrame([{"CPF": valid_cpf(1)}])
    data = {"base": xlsx_upload(base, "base.xlsx"), "lista": xlsx_upload(lista, "lista.xlsx")}
    resp = client.post("/api/process_inativacao", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert os.listdir(settings.UPLOAD_FOLDER) == []
