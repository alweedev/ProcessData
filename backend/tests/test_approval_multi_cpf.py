import pandas as pd
import pytest
from _helpers import format_cpf, valid_cpf

from backend.services.approval_service import ApprovalService

A, B, C, D = (valid_cpf(i) for i in (1, 2, 3, 4))


def _df(*rows):
    return pd.DataFrame(list(rows)).fillna("").astype(str)


def _cols(df):
    return ApprovalService.detect_approval_columns(df)


def _linha(aid, *aprovadores, por="CCEMPRESA", cpf="", segundo=""):
    linha = {"AprovacaoId": aid, "AprovacaoPor": por, "CPF": cpf}
    for i, login in enumerate(aprovadores, 1):
        linha[f"LoginAprovador_{i}"] = login
    if segundo:
        linha["LoginAprovador_SEGUNDO_NIVEL"] = segundo
    return linha


def test_detecta_coluna_cpf_do_viajante():
    assert _cols(_df(_linha("S1", A)))["traveler_cpf_col"] == "CPF"


def test_find_traveler_structures_so_conta_viajante_e_aceita_pontuacao():
    df = _df(
        _linha("S1", B, por="VIAJANTE", cpf=format_cpf(A)),
        _linha("S2", B, por="VIAJANTE", cpf=B),
        _linha("S3", B, por="CCEMPRESA", cpf=A),
    )
    assert ApprovalService.find_traveler_structures(df, {A, B}, _cols(df)) == {A: {"S1"}, B: {"S2"}}


def test_find_traveler_structures_sem_coluna_cpf_levanta_erro():
    df = _df(_linha("S1", A)).drop(columns=["CPF"])
    with pytest.raises(ValueError, match="CPF do viajante"):
        ApprovalService.find_traveler_structures(df, {A}, _cols(df))


def test_find_approver_structures_devolve_posicoes_e_segundo_nivel():
    df = _df(_linha("S1", A, B), _linha("S2", C, A, segundo=A))
    achados = ApprovalService.find_approver_structures(df, {A, B}, _cols(df))
    assert achados[A] == {
        "S1": {"posicoes": [1], "segundoNivel": False},
        "S2": {"posicoes": [2], "segundoNivel": True},
    }
    assert achados[B] == {"S1": {"posicoes": [2], "segundoNivel": False}}


def test_remove_varios_cpfs_compacta_uma_vez():
    df = _df(_linha("S1", A, B, C, D))
    out, stats = ApprovalService.remove_cpfs_and_compact(df, {A, C}, _cols(df), {"S1"}, True)
    assert out.loc[0, ["LoginAprovador_1", "LoginAprovador_2", "LoginAprovador_3", "LoginAprovador_4"]].tolist() == [
        B,
        D,
        "",
        "",
    ]
    assert stats["structures_updated"] == 1
    assert stats["occurrences_removed"] == 2


def test_remove_varios_cpfs_equivale_a_remocoes_sequenciais():
    df = _df(_linha("S1", A, B, C), _linha("S2", B, A), _linha("S3", C), _linha("S4", D, A, segundo=B))
    cols = _cols(df)
    ids = {"S1", "S2", "S3", "S4"}
    de_uma_vez, _ = ApprovalService.remove_cpfs_and_compact(df, {A, B}, cols, ids, True)
    passo, _ = ApprovalService.remove_cpf_and_compact(df, A, cols, ids, True)
    em_sequencia, _ = ApprovalService.remove_cpf_and_compact(passo, B, cols, ids, True)
    assert de_uma_vez.equals(em_sequencia)


def test_promove_segundo_nivel_quando_o_primeiro_esvazia():
    df = _df(_linha("S1", A, segundo=C))
    out, stats = ApprovalService.remove_cpfs_and_compact(df, {A}, _cols(df), {"S1"}, True)
    assert out.loc[0, "LoginAprovador_1"] == C
    assert out.loc[0, "LoginAprovador_SEGUNDO_NIVEL"] == ""
    assert stats["promotions"] == 1


def test_segundo_nivel_de_outro_cpf_da_lista_nao_e_promovido():
    df = _df(_linha("S1", A, segundo=B))
    out, stats = ApprovalService.remove_cpfs_and_compact(df, {A, B}, _cols(df), {"S1"}, True)
    assert out.loc[0, "LoginAprovador_1"] == ""
    assert out.loc[0, "LoginAprovador_SEGUNDO_NIVEL"] == ""
    assert stats["promotions"] == 0


def test_structures_left_without_approvers_considera_todos_os_cpfs():
    df = _df(_linha("S1", A, B), _linha("S2", A, C))
    orfas = ApprovalService.structures_left_without_approvers(df, {A, B}, _cols(df), {"S1", "S2"}, True)
    assert [o["aprovacaoId"] for o in orfas] == ["S1"]


def test_delete_structures_marca_operacao_sem_mexer_na_entrada():
    df = _df(_linha("S1", A, por="VIAJANTE", cpf=A), _linha("S2", B))
    saida = ApprovalService.delete_structures(df, {"S1"}, _cols(df))
    assert saida["AprovacaoId"].tolist() == ["S1"]
    assert saida["Operacao"].tolist() == ["DELETE"]
    assert "Operacao" not in df.columns


def test_nao_altera_a_entrada():
    df = _df(_linha("S1", A, B), _linha("S2", C, segundo=A))
    antes = df.copy()
    ApprovalService.remove_cpfs_and_compact(df, {A}, _cols(df), {"S1", "S2"}, True)
    assert df.equals(antes)
