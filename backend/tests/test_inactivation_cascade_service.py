import pytest
from _cascade_fixtures import USR_A, USR_B, USR_C, A, B, C, D, cad, est, viajante
from _helpers import valid_cpf

from backend.core.config import settings
from backend.services.inactivation_cascade_service import (
    ALERTA_SEM_CPF,
    InactivationCascadeService,
    InativacaoError,
)


def _analisar(cadastro, estruturas, itens, **kw):
    return InactivationCascadeService.analisar(cadastro, estruturas, itens, **kw)


def _usuario(analise, indice=0):
    return analise.payload["usuarios"][indice]


def test_estrutura_do_viajante_e_marcada_para_exclusao():
    a = _analisar(cad(USR_A), est(viajante("S1", A, B)), [A])
    u = _usuario(a)
    assert u["situacao"] == "EXECUTAVEL"
    assert u["estruturasViajante"] == ["S1"]
    assert a.payload["resumo"]["estruturasExcluidas"] == 1
    assert a.excluidas == {"S1"}


def test_aprovador_com_substitutos_vira_compactacao():
    a = _analisar(cad(USR_A), est(viajante("S1", C, B, A)), [A])
    u = _usuario(a)
    assert u["estruturasViajante"] == []
    assert u["comoAprovador"] == [{"aprovacaoId": "S1", "posicoes": [2], "segundoNivel": False, "acao": "COMPACTACAO"}]
    assert a.payload["resumo"]["estruturasCompactadas"] == 1


def test_aprovador_unico_vira_estrutura_orfa():
    a = _analisar(cad(USR_A), est(viajante("S1", C, A)), [A])
    assert _usuario(a)["comoAprovador"][0]["acao"] == "ORFA"
    assert a.payload["resumo"]["estruturasOrfas"] == 1
    assert a.orfas == {"S1"}


def test_dois_da_lista_que_aprovam_uma_estrutura_deixam_ela_orfa_para_ambos():
    a = _analisar(cad(USR_A, USR_B), est(viajante("S1", C, A, B)), [A, B])
    assert {u["cpf"]: u["comoAprovador"][0]["acao"] for u in a.payload["usuarios"]} == {A: "ORFA", B: "ORFA"}
    assert a.payload["resumo"]["estruturasOrfas"] == 1


def test_estrutura_excluida_nunca_conta_como_orfa():
    a = _analisar(cad(USR_A), est(viajante("S1", A, A)), [A])
    assert _usuario(a)["estruturasViajante"] == ["S1"]
    assert _usuario(a)["comoAprovador"] == []
    assert a.payload["resumo"]["estruturasOrfas"] == 0


def test_segundo_nivel_e_removido_e_conta_como_compactacao():
    a = _analisar(cad(USR_A), est(viajante("S1", C, B, segundo=A)), [A])
    assert _usuario(a)["comoAprovador"] == [
        {"aprovacaoId": "S1", "posicoes": [], "segundoNivel": True, "acao": "COMPACTACAO"}
    ]


def test_usuario_sem_cpf_nao_e_executavel_e_traz_o_alerta_exato():
    a = _analisar(cad(("", "Sem Cpf Silva", "s@x.com", "ATIVO")), est(viajante("S1", C, B)), ["Sem Cpf Silva"])
    u = _usuario(a)
    assert u["situacao"] == "SEM_CPF"
    assert u["alerta"] == ALERTA_SEM_CPF
    assert a.cpfs == frozenset()


def test_usuario_ja_inativo_e_nao_localizado():
    a = _analisar(cad((A, "Ana Souza", "ana@x.com", "INATIVO")), est(viajante("S1", C, B)), [A, valid_cpf(9)])
    assert [u["situacao"] for u in a.payload["usuarios"]] == ["JA_INATIVO", "NAO_LOCALIZADO"]
    assert a.cpfs == frozenset()


def test_homonimos_exigem_selecao_explicita():
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    estruturas = est(viajante("S9", C, D))
    pendente = _analisar(cadastro, estruturas, ["Joao Silva"])
    u = _usuario(pendente)
    assert u["situacao"] == "PENDENTE_SELECAO"
    assert {c["cpf"] for c in u["candidatos"]} == {A, B}
    assert pendente.cpfs == frozenset()

    escolhido = _analisar(cadastro, estruturas, ["Joao Silva"], selecionados=[A])
    assert [u["situacao"] for u in escolhido.payload["usuarios"]] == ["EXECUTAVEL"]
    assert escolhido.cpfs == {A}


def test_cpf_com_pontuacao_e_zero_a_esquerda_perdido():
    cadastro = cad(("1234567890", "Ana Souza", "ana@x.com", "ATIVO"))
    a = _analisar(cadastro, est(viajante("S1", C, "012.345.678-90", B)), ["01234567890"])
    assert _usuario(a)["cpf"] == "01234567890"
    assert _usuario(a)["comoAprovador"][0]["aprovacaoId"] == "S1"


def test_cpf_repetido_na_lista_entra_uma_vez():
    a = _analisar(cad(USR_A), est(viajante("S1", C, B)), [A, A])
    assert len(a.payload["usuarios"]) == 1
    assert a.payload["resumo"]["duplicados"] == [A]


def test_nao_altera_as_entradas():
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, A))
    antes_c, antes_e = cadastro.copy(), estruturas.copy()
    _analisar(cadastro, estruturas, [A])
    assert cadastro.equals(antes_c)
    assert estruturas.equals(antes_e)


def test_impressao_digital_estavel_e_sensivel_ao_impacto():
    cadastro = cad(USR_A)
    base = est(viajante("S1", C, A, B))
    primeira = _analisar(cadastro, base, [A]).payload["impressaoDigital"]
    assert primeira == _analisar(cadastro, base, [A]).payload["impressaoDigital"]
    outra = est(viajante("S1", C, A))  # agora A é o único aprovador
    assert primeira != _analisar(cadastro, outra, [A]).payload["impressaoDigital"]


def test_lista_vazia():
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), est(viajante("S1", C, B)), ["  "])
    assert erro.value.code == "LISTA_VAZIA"


def test_lista_grande(monkeypatch):
    monkeypatch.setattr(settings, "MAX_INATIVACAO_ITENS", 2, raising=False)
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), est(viajante("S1", C, B)), [A, B, C])
    assert erro.value.code == "LISTA_GRANDE"


def test_base_de_estruturas_sem_coluna_cpf():
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), est(viajante("S1", C, B)).drop(columns=["CPF"]), [A])
    assert erro.value.code == "BASE_SEM_COLUNA"
    assert "CPF" in erro.value.message


def test_base_de_cadastro_sem_coluna_cpf():
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A).drop(columns=["CPF"]), est(viajante("S1", C, B)), ["Ana Souza"])
    assert erro.value.code == "BASE_SEM_COLUNA"


def test_base_sem_linha_viajante_nao_exige_cpf_do_viajante():
    df_cad = cad(USR_A)
    df_est = est({"AprovacaoId": "X1", "AprovacaoPor": "CCEMPRESA", "LoginAprovador_1": "999.999.999-99"})
    # Não deve levantar InativacaoError("BASE_SEM_COLUNA", ...): a base não usa estrutura por viajante.
    analise = InactivationCascadeService.analisar(df_cad, df_est, [A])
    assert analise.payload["usuarios"][0]["situacao"] == "EXECUTAVEL"


def test_usuario_fora_da_lista_nao_e_tocado():
    a = _analisar(cad(USR_A, USR_C), est(viajante("S1", D, A, C)), [A])
    assert a.cpfs == {A}
    assert a.orfas == frozenset()


def test_homonimo_digitado_por_nome_e_outro_por_cpf_nao_entra_sem_escolha():
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    a = _analisar(cadastro, est(viajante("S9", C, D)), ["Joao Silva", A])
    situacoes = {u["cpf"]: u["situacao"] for u in a.payload["usuarios"] if u["cpf"]}
    assert situacoes == {A: "EXECUTAVEL"}
    pendente = [u for u in a.payload["usuarios"] if u["situacao"] == "PENDENTE_SELECAO"]
    assert len(pendente) == 1
    assert {c["cpf"] for c in pendente[0]["candidatos"]} == {B}
    assert a.cpfs == {A}


def test_homonimo_digitado_por_nome_e_outro_por_email_nao_entra_sem_escolha():
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    a = _analisar(cadastro, est(viajante("S9", C, D)), ["Joao Silva", "j1@x.com"])
    assert [u["situacao"] for u in a.payload["usuarios"]].count("PENDENTE_SELECAO") == 1
    assert [u["cpf"] for u in a.payload["usuarios"] if u["situacao"] == "EXECUTAVEL"] == [A]
    assert a.cpfs == {A}


def test_homonimo_digitado_por_nome_entra_quando_escolhido():
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    a = _analisar(cadastro, est(viajante("S9", C, D)), ["Joao Silva", A], selecionados=[B])
    assert {u["cpf"]: u["situacao"] for u in a.payload["usuarios"]} == {A: "EXECUTAVEL", B: "EXECUTAVEL"}
    assert a.cpfs == {A, B}


def test_status_em_branco_com_coluna_de_status_nao_e_executavel():
    a = _analisar(cad((A, "Ana Souza", "ana@x.com", "")), est(viajante("S1", A, B)), [A])
    u = _usuario(a)
    assert u["situacao"] == "JA_INATIVO"
    assert u["alerta"] == "Usuário não está ATIVO no cadastro (Status: '')."
    assert a.cpfs == frozenset()


def test_cadastro_sem_coluna_de_status_continua_executavel():
    cadastro = cad(USR_A).drop(columns=["Status"])
    a = _analisar(cadastro, est(viajante("S1", A, B)), [A])
    assert _usuario(a)["situacao"] == "EXECUTAVEL"
    assert a.cpfs == {A}


def _linha_ccempresa(aid, *aprovadores):
    linha = {"AprovacaoId": aid, "AprovacaoPor": "CCEMPRESA", "CPF": "", "NomeViajante": ""}
    for i, login in enumerate(aprovadores, 1):
        linha[f"LoginAprovador_{i}"] = login
    return linha


def test_id_reaproveitado_por_outro_viajante_e_recusado():
    estruturas = est(viajante("S1", A, B), viajante("S1", C, B))
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), estruturas, [A])
    assert erro.value.code == "ESTRUTURA_COMPARTILHADA"
    assert erro.value.status == 400
    assert "S1" in erro.value.message


def test_id_que_mistura_linha_ccempresa_e_recusado():
    estruturas = est(viajante("S1", A, B), _linha_ccempresa("S1", B))
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), estruturas, [A])
    assert erro.value.code == "ESTRUTURA_COMPARTILHADA"


def test_id_com_linha_de_viajante_sem_cpf_e_recusado_e_lista_os_ids_ordenados():
    estruturas = est(viajante("S2", A, B), viajante("S2", "", B), viajante("S1", A, B), viajante("S1", C, B))
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), estruturas, [A])
    assert "S1, S2" in erro.value.message


def test_id_com_varias_linhas_do_mesmo_viajante_e_permitido():
    estruturas = est(viajante("S1", A, B), viajante("S1", A, C), viajante("S2", C, B))
    a = _analisar(cad(USR_A), estruturas, [A])
    assert a.excluidas == {"S1"}


def test_id_compartilhado_so_importa_quando_o_dono_esta_no_conjunto_executavel():
    estruturas = est(viajante("S1", A, B), viajante("S1", C, B))
    assert _analisar(cad(USR_A, USR_C), estruturas, [A, C]).excluidas == {"S1"}
    assert _analisar(cad(USR_A), est(viajante("S1", C, B), viajante("S1", C, D)), [A]).excluidas == frozenset()


def test_itens_que_nao_sao_cpf_email_nem_nome_completo_viram_nao_localizado():
    a = _analisar(cad(USR_A), est(viajante("S1", A, B)), ["1234567890", "Maria", A])
    usuarios = a.payload["usuarios"]
    assert [u["situacao"] for u in usuarios] == ["EXECUTAVEL", "NAO_LOCALIZADO", "NAO_LOCALIZADO"]
    ignorados = usuarios[1:]
    assert [u["nome"] for u in ignorados] == ["1234567890", "Maria"]
    for u in ignorados:
        assert u["cpf"] is None
        assert u["cpfMascarado"] == ""
        assert u["alerta"] == "Item não reconhecido como CPF (11 dígitos), e-mail ou nome completo."
    assert a.cpfs == {A}
