import pandas as pd

from backend.services.argo_schema_validator import ArgoSchemaValidator


def test_aceita_planilha_valida():
    df = pd.DataFrame([
        {"Operacao": "DELETE", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-11", "Status": ""},
    ])
    assert ArgoSchemaValidator.validar(df) == []


def test_rejeita_operacao_vazia():
    df = pd.DataFrame([
        {"Operacao": "", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-11", "Status": ""},
    ])
    erros = ArgoSchemaValidator.validar(df)
    assert any("Operacao" in e for e in erros)


def test_rejeita_login_aprovador_maior_que_12_caracteres():
    df = pd.DataFrame([
        {"Operacao": "UPDATE", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-111-extra", "Status": ""},
    ])
    erros = ArgoSchemaValidator.validar(df)
    assert any("LoginAprovador_1" in e for e in erros)


def test_rejeita_status_preenchido():
    df = pd.DataFrame([
        {"Operacao": "DELETE", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-11", "Status": "Ativo"},
    ])
    erros = ArgoSchemaValidator.validar(df)
    assert any("Status" in e for e in erros)
