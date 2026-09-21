"""Conferência de nomes: separação, limite de 20 caracteres sem corte, correções do usuário e aprendizado."""

import io
import json
import os
import time

import pandas as pd
import pytest
from _helpers import valid_cpf, xlsx_upload

from backend.core.config import settings
from backend.services.processing_service import ProcessingService
from backend.shared.name_splitter import MAX_FIELD_LEN


def _row(i=1, nome="Maria Clara da Silva Santos", **override):
    row = {
        "CPF": valid_cpf(i),
        "NOME COMPLETO": nome,
        "E-MAIL": f"p{i}@x.com",
        "EMPRESA": "Empresa A",
        "Centro de custo": "CC1",
        "Descrição Centro de Custo": "ADMINISTRATIVO",
        "TELEFONE": "11999990000",
        "Data de Nascimento": "12/05/1990",
        "SOLICITANTE? (S/N)": "S",
    }
    row.update(override)
    return row


def _write(rows, path):
    pd.DataFrame(rows).to_excel(path, index=False)
    return str(path)


def _run(rows, tmp_path, **kwargs):
    return ProcessingService.process_records_from_files([_write(rows, tmp_path / "in.xlsx")], **kwargs)


def _override(nome_completo, nome, sobrenome, editado=False):
    return {"nome_completo": nome_completo, "nome": nome, "sobrenome": sobrenome, "editado": editado}


LONG_SURNAME = "Fernandes de Albuquerque Cavalcanti"  # 35 caracteres: não cabe em 20
# Nenhuma palavra no vocabulário e cabe em 20: a divisão é ambígua (vai para conferência), mas não estoura.
DUVIDOSO = "Xanto Zeko Quim Brag"
DUVIDOSO_COMPLETO = "XANTO ZEKO QUIM BRAG"


# ------------------------------------------------------------------ pipeline


def test_nome_e_sobrenome_saem_como_no_documento(tmp_path):
    _errors, df = _run([_row(1)], tmp_path)

    assert df["Nome"].tolist() == ["MARIA CLARA"]
    assert df["SobreNome"].tolist() == ["DA SILVA SANTOS"]
    assert df["NomeCompleto"].tolist() == ["MARIA CLARA DA SILVA SANTOS"]
    assert df.attrs["name_review"] == []  # divisão segura: ninguém precisa conferir


def test_nome_que_nao_cabe_em_20_nao_e_cortado_e_vai_para_conferencia(tmp_path):
    _errors, df = _run([_row(1, f"Anna Luiza {LONG_SURNAME}")], tmp_path)

    assert df["SobreNome"].tolist() == ["FERNANDES DE ALBUQUERQUE CAVALCANTI"]  # inteiro, como no documento
    [item] = df.attrs["name_review"]
    assert item["estouro"] == {"nome": False, "sobrenome": True}
    assert item["sugestao"] == {"nome": "ANNA LUIZA", "sobrenome": "F DE A CAVALCANTI"}
    assert len(item["sugestao"]["sobrenome"]) <= MAX_FIELD_LEN
    assert item["label"] == "Linha 2" and item["key"] == "1:2"


def test_nome_duvidoso_entra_na_conferencia_com_o_motivo_e_o_seguro_nao(tmp_path):
    rows = [_row(1), _row(2, DUVIDOSO), _row(3, "Ana Paula")]

    _errors, df = _run(rows, tmp_path)

    review = {item["key"]: item for item in df.attrs["name_review"]}
    assert set(review) == {"1:3", "1:4"}  # a linha 2 (Maria Clara da Silva Santos) segue sozinha
    assert review["1:3"]["confianca"] == "baixa"
    assert any("ambígua" in m for m in review["1:3"]["motivos"])
    assert review["1:3"]["estouro"] == {"nome": False, "sobrenome": False}
    assert review["1:3"]["sugestao"] is None
    assert any("sem sobrenome" in m for m in review["1:4"]["motivos"])


def test_nome_e_sobrenome_da_ficha_sao_ignorados_quando_ha_nome_completo(tmp_path):
    row = _row(1, "Carlos Eduardo Lima", NOME="Cadu", **{"SOBRENOME (limite 20 caracteres)": "Limão"})

    _errors, df = _run([row], tmp_path)

    assert (df["Nome"].tolist(), df["SobreNome"].tolist()) == (["CARLOS EDUARDO"], ["LIMA"])


def test_sem_nome_completo_valem_o_nome_e_o_sobrenome_da_ficha(tmp_path):
    row = _row(1, "", NOME="Carlos", **{"SOBRENOME (limite 20 caracteres)": "Pereira"})

    errors, df = _run([row], tmp_path)

    assert (df["Nome"].tolist(), df["SobreNome"].tolist()) == (["CARLOS"], ["PEREIRA"])
    assert errors[0] == "Campo obrigatório em branco: Nome completo"
    assert df.attrs["name_review"] == []  # sem nome completo não há o que conferir (a linha já é inválida)


def test_chave_da_conferencia_identifica_arquivo_e_linha(tmp_path):
    a = _write([_row(1)], tmp_path / "a.xlsx")
    b = _write([_row(2), _row(3, DUVIDOSO)], tmp_path / "b.xlsx")

    _errors, df = ProcessingService.process_records_from_files([a, b])

    [item] = df.attrs["name_review"]
    assert item["key"] == "2:3" and item["label"] == "Arquivo 2 · linha 3"


# --------------------------------------------------------------- correções


def test_correcao_do_usuario_vale_mais_que_a_sugestao(tmp_path):
    overrides = {"1:2": _override(DUVIDOSO_COMPLETO, "XANTO", "ZEKO QUIM BRAG", True)}

    _errors, df = _run([_row(1, DUVIDOSO)], tmp_path, name_overrides=overrides)

    assert (df["Nome"].tolist(), df["SobreNome"].tolist()) == (["XANTO"], ["ZEKO QUIM BRAG"])
    assert df.attrs["name_review"] == []  # já decidido: não volta para a conferência
    assert df.attrs["name_overrides_applied"] == [{"nome": "XANTO", "sobrenome": "ZEKO QUIM BRAG", "editado": True}]


def test_correcao_de_outra_planilha_nao_e_herdada(tmp_path):
    # mesma linha, mas o nome completo mudou (planilha trocada): a correção antiga não vale
    overrides = {"1:2": _override(DUVIDOSO_COMPLETO, "XANTO", "ZEKO QUIM BRAG")}

    _errors, df = _run([_row(1, "Ana Silva")], tmp_path, name_overrides=overrides)

    assert (df["Nome"].tolist(), df["SobreNome"].tolist()) == (["ANA"], ["SILVA"])
    assert df.attrs["name_overrides_applied"] == []


def test_correcao_que_ainda_passa_de_20_continua_na_conferencia(tmp_path):
    completo = f"Anna Luiza {LONG_SURNAME}".upper()
    overrides = {"1:2": _override(completo, "ANNA LUIZA", "FERNANDES DE ALBUQUERQUE CAVALCANTI")}

    _errors, df = _run([_row(1, f"Anna Luiza {LONG_SURNAME}")], tmp_path, name_overrides=overrides)

    [item] = df.attrs["name_review"]
    assert item["estouro"]["sobrenome"] is True  # aceitar não resolve: tem que caber


def test_correcao_que_cabe_resolve_o_estouro(tmp_path):
    completo = f"Anna Luiza {LONG_SURNAME}".upper()
    overrides = {"1:2": _override(completo, "ANNA LUIZA", "F DE A CAVALCANTI", True)}

    _errors, df = _run([_row(1, f"Anna Luiza {LONG_SURNAME}")], tmp_path, name_overrides=overrides)

    assert df["SobreNome"].tolist() == ["F DE A CAVALCANTI"]
    assert df.attrs["name_review"] == []
    assert df["NomeCompleto"].tolist() == [completo]  # o nome completo do documento não muda


# --------------------------------------------------------------------- API


def _post(client, route, rows, **form):
    data = {"files[]": xlsx_upload(pd.DataFrame(rows), "fichas.xlsx"), "login_choice": "CPF", "fluxo": "SELF", **form}
    return client.post(route, data=data, content_type="multipart/form-data")


def test_validacao_devolve_os_nomes_para_conferir(client):
    resp = _post(client, "/api/analysis/summary", [_row(1), _row(2, DUVIDOSO)])

    body = resp.get_json()
    assert resp.status_code == 200
    assert [i["key"] for i in body["name_review"]] == ["1:3"]
    assert body["name_review"][0]["nome_completo"] == DUVIDOSO_COMPLETO


def test_historico_nao_guarda_nome_de_passageiro(client):
    _post(client, "/api/analysis/summary", [_row(1, DUVIDOSO)])

    with open(settings.HISTORY_LOG_FILE, encoding="utf-8") as fh:
        historico = fh.read()
    assert "names_to_review" in historico
    assert "XANTO" not in historico.upper() and "ZEKO" not in historico.upper()


def test_geracao_recusa_nome_acima_de_20_com_a_linha_no_aviso(client):
    resp = _post(client, "/api/process_cadastro", [_row(1), _row(2, f"Anna Luiza {LONG_SURNAME}")])

    assert resp.status_code == 400
    assert resp.get_json()["error"] == (
        "1 nome com mais de 20 caracteres em Nome ou Sobrenome (Linha 3). Ajuste na conferência de nomes antes de gerar."
    )


def test_geracao_com_a_correcao_que_cabe_passa_e_sai_o_nome_ajustado(client):
    completo = f"Anna Luiza {LONG_SURNAME}".upper()
    overrides = {"1:2": _override(completo, "ANNA LUIZA", "F DE A CAVALCANTI", True)}

    resp = _post(
        client, "/api/process_cadastro", [_row(1, f"Anna Luiza {LONG_SURNAME}")], name_overrides=json.dumps(overrides)
    )

    assert resp.status_code == 200
    out = pd.read_excel(io.BytesIO(resp.data), dtype=str).fillna("")
    assert (out["Nome"][0], out["SobreNome"][0]) == ("ANNA LUIZA", "F DE A CAVALCANTI")


def test_geracao_com_nome_duvidoso_nao_e_recusada_no_servidor(client):
    # a exigência de "aceitar" é da tela; o servidor só barra o que a plataforma não aceita (mais de 20)
    resp = _post(client, "/api/process_cadastro", [_row(1, DUVIDOSO)])

    assert resp.status_code == 200


def test_aprende_so_com_o_que_o_usuario_confirmou_e_a_correcao_pesa_mais(client):
    aceito = _override("BELTRANO ZEFERINO QUINTELA BRAGANCA", "BELTRANO ZEFERINO", "QUINTELA BRAGANCA")
    corrigido = _override("ANA PAULA MARIA XAVIER", "ANA PAULA", "MARIA XAVIER", True)
    rows = [_row(1, "Beltrano Zeferino Quintela Bragança"), _row(2, "Ana Paula Maria Xavier")]

    resp = _post(client, "/api/process_cadastro", rows, name_overrides=json.dumps({"1:2": aceito, "1:3": corrigido}))

    assert resp.status_code == 200
    with open(settings.NAME_VOCAB_FILE, encoding="utf-8") as fh:
        saved = json.load(fh)
    assert saved["given"]["BELTRANO"] == 1 and saved["given"]["ZEFERINO"] == 1  # aceitou: peso 1
    assert saved["surname"]["QUINTELA"] == 1
    assert saved["given"]["PAULA"] == 3 and saved["surname"]["MARIA"] == 3  # corrigiu: peso 3


def test_sem_correcoes_nada_e_aprendido(client):
    resp = _post(client, "/api/process_cadastro", [_row(1, DUVIDOSO)])

    assert resp.status_code == 200
    assert not os.path.exists(settings.NAME_VOCAB_FILE)  # gerou sem conferir nada: o vocabulário não aprende


def test_o_aprendizado_muda_a_proxima_validacao(client):
    nome = "Beltrano Zeferino Quintela Bragança"
    completo = "BELTRANO ZEFERINO QUINTELA BRAGANCA"
    assert len(_post(client, "/api/analysis/summary", [_row(1, nome)]).get_json()["name_review"]) == 1

    corrigido = _override(completo, "BELTRANO ZEFERINO", "QUINTELA BRAGANCA", True)
    for _ in range(2):  # a mesma correção duas vezes ensina o vocabulário
        _post(client, "/api/process_cadastro", [_row(1, nome)], name_overrides=json.dumps({"1:2": corrigido}))

    depois = _post(client, "/api/analysis/summary", [_row(1, nome)]).get_json()
    assert depois["name_review"] == []  # agora a divisão é segura e ninguém precisa conferir


@pytest.mark.parametrize(
    "corpo",
    [
        "isso não é json",
        json.dumps(["lista"]),
        json.dumps({"1:2": "texto"}),
        json.dumps({"1:2": {"nome_completo": "A", "nome": 1, "sobrenome": "B"}}),
        json.dumps({"1:2": {"nome_completo": "A", "nome": "X" * 500, "sobrenome": "B"}}),
        json.dumps({"k" * 100: _override("A", "B", "C")}),
    ],
)
def test_correcoes_mal_formadas_sao_recusadas(client, corpo):
    resp = _post(client, "/api/process_cadastro", [_row(1)], name_overrides=corpo)

    assert resp.status_code == 400
    assert resp.get_json()["error"] == "Correções de nomes inválidas."


# ---------------------------------------------------------------- volume


def test_20_mil_nomes_diferentes_separam_em_tempo_razoavel(tmp_path):
    rows = [_row(i, f"Pessoa{i} Nome{i} Sobrenome{i}") for i in range(1, 3001)]
    path = _write(rows, tmp_path / "grande.xlsx")

    start = time.time()
    ProcessingService.process_records_from_files([path])

    assert time.time() - start < 20  # folga grande (leva ~1 s): só pega regressão para algo quadrático
