import io
import json
import os

import pandas as pd
from _cascade_fixtures import USR_A, A, B, C, cad, est, viajante
from _helpers import xlsx_upload

from backend.core.config import settings
from backend.services.audit_service import AuditService
from backend.services.inactivation_cascade_service import InactivationCascadeService

ROTA = "/api/inativacao/analisar"


def _post(client, cadastro=None, estruturas=None, **campos):
    data = dict(campos)
    if cadastro is not None:
        data["cadastro"] = xlsx_upload(cadastro, "cadastro.xlsx")
    if estruturas is not None:
        data["estruturas"] = xlsx_upload(estruturas, "estruturas.xlsx")
    return client.post(ROTA, data=data, content_type="multipart/form-data")


def _bases():
    return cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))


def test_analisa_por_itens_json(client):
    resp = _post(client, *_bases(), itens=json.dumps([A]))
    assert resp.status_code == 200, resp.get_data(as_text=True)
    corpo = resp.get_json()
    assert set(corpo) == {"usuarios", "resumo", "impressaoDigital"}
    assert corpo["usuarios"][0]["situacao"] == "EXECUTAVEL"
    assert corpo["usuarios"][0]["estruturasViajante"] == ["S1"]
    assert corpo["resumo"]["estruturasCompactadas"] == 1


def test_analisa_por_lista_text(client):
    resp = _post(client, *_bases(), lista_text=f"{A}\n")
    assert resp.status_code == 200
    assert resp.get_json()["resumo"]["executaveis"] == 1


def test_analisa_por_lista_em_arquivo(client):
    lista = xlsx_upload(pd.DataFrame([{"CPF": A}]), "lista.xlsx")
    resp = _post(client, *_bases(), lista=lista)
    assert resp.status_code == 200
    assert resp.get_json()["resumo"]["executaveis"] == 1


def test_selecionados_resolvem_homonimos(client):
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    estruturas = est(viajante("S9", C, A))
    pendente = _post(client, cadastro, estruturas, itens=json.dumps(["Joao Silva"])).get_json()
    assert pendente["usuarios"][0]["situacao"] == "PENDENTE_SELECAO"
    escolhido = _post(client, cadastro, estruturas, itens=json.dumps(["Joao Silva"]), selecionados=json.dumps([B]))
    assert escolhido.get_json()["usuarios"][0]["situacao"] == "EXECUTAVEL"


def test_sem_a_base_de_estruturas(client):
    resp = _post(client, cad(USR_A), None, itens=json.dumps([A]))
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "BASE_AUSENTE"


def test_arquivo_que_nao_e_excel(client):
    data = {
        "cadastro": (io.BytesIO(b"isto nao e um excel"), "cadastro.xlsx"),
        "estruturas": xlsx_upload(_bases()[1], "estruturas.xlsx"),
        "itens": json.dumps([A]),
    }
    resp = client.post(ROTA, data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "ARQUIVO_INVALIDO"


def test_lista_vazia(client):
    resp = _post(client, *_bases())
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "LISTA_VAZIA"


def test_base_de_estruturas_sem_coluna_cpf(client):
    resp = _post(client, cad(USR_A), _bases()[1].drop(columns=["CPF"]), itens=json.dumps([A]))
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "BASE_SEM_COLUNA"


def test_arquivo_acima_do_teto_da_rota(client, monkeypatch):
    monkeypatch.setattr(settings, "INATIVACAO_MAX_CONTENT_LENGTH", 200, raising=False)
    resp = _post(client, *_bases(), itens=json.dumps([A]))
    assert resp.status_code == 413
    assert resp.get_json()["code"] == "ARQUIVO_GRANDE"


def test_apaga_os_temporarios_no_sucesso_e_no_erro(client):
    _post(client, *_bases(), itens=json.dumps([A]))
    _post(client, *_bases())
    assert os.listdir(settings.UPLOAD_FOLDER) == []


def test_auditoria_nao_guarda_dado_pessoal(client):
    assert _post(client, *_bases(), itens=json.dumps([A])).status_code == 200
    eventos = AuditService.list_events()
    assert any(e["event_type"] == "inativacao_analise" and e["status"] == "success" for e in eventos)
    texto = json.dumps(eventos, ensure_ascii=False)
    for pessoal in (A, "Ana Souza", "ana@x.com"):
        assert pessoal not in texto
    assert "***." in texto


def test_auditoria_mascara_cpf_duplicado(client):
    resp = _post(client, *_bases(), itens=json.dumps([A, A]))
    assert resp.status_code == 200
    assert resp.get_json()["resumo"]["duplicados"] == [A]  # a resposta ao operador segue completa
    texto = json.dumps(AuditService.list_events(), ensure_ascii=False)
    assert A not in texto
    assert "***." in texto


def test_erro_interno_vira_500_sem_vazar_detalhe(client, monkeypatch):
    def boom(*args, **kwargs):
        raise RuntimeError("segredo interno")

    monkeypatch.setattr(InactivationCascadeService, "analisar", boom)
    resp = _post(client, *_bases(), itens=json.dumps([A]))
    assert resp.status_code == 500
    assert resp.get_json()["code"] == "ERRO_INTERNO"
    assert "segredo" not in resp.get_data(as_text=True)


def test_lista_ilegivel_vira_400_arquivo_invalido_e_limpa_temporarios(client):
    data = {
        "cadastro": xlsx_upload(_bases()[0], "cadastro.xlsx"),
        "estruturas": xlsx_upload(_bases()[1], "estruturas.xlsx"),
        "lista": (io.BytesIO(b"CPF\n123\n"), "lista.xls"),
    }
    resp = client.post(ROTA, data=data, content_type="multipart/form-data")
    assert resp.status_code == 400, resp.get_data(as_text=True)
    assert resp.get_json()["code"] == "ARQUIVO_INVALIDO"
    assert os.listdir(settings.UPLOAD_FOLDER) == []


def test_teto_da_rota_acima_dos_16_mb_globais(client):
    grande = io.BytesIO(b"0" * (17 * 1024 * 1024))
    resp = client.post(ROTA, data={"lixo": (grande, "lixo.bin")}, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "BASE_AUSENTE"
