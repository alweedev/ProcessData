import io
import zipfile

import pytest
from _cascade_fixtures import USR_A, A, B, C, D, cad, est, viajante

from backend.services.export_service import ExportService
from backend.services.inactivation_cascade_service import InactivationCascadeService, InativacaoError


def _digital(cadastro, estruturas, cpfs):
    return InactivationCascadeService.analisar(cadastro, estruturas, cpfs).payload["impressaoDigital"]


def _executar(cadastro, estruturas, cpfs, **kw):
    return InactivationCascadeService.executar(cadastro, estruturas, cpfs, _digital(cadastro, estruturas, cpfs), **kw)


def test_gera_ficha_delete_e_estruturas_com_delete_e_update():
    estruturas = est(viajante("S1", A, B), viajante("S2", C, B, A))
    r = _executar(cad(USR_A), estruturas, [A])
    assert r.ficha["Operacao"].tolist() == ["DELETE"]
    assert dict(zip(r.estruturas["AprovacaoId"], r.estruturas["Operacao"])) == {"S1": "DELETE", "S2": "UPDATE"}
    s2 = r.estruturas[r.estruturas["AprovacaoId"] == "S2"].iloc[0]
    assert (s2["LoginAprovador_1"], s2["LoginAprovador_2"]) == (B, "")
    assert r.estruturas.columns[0] == "Operacao"
    assert r.resumo == {
        "usuariosInativados": 1,
        "estruturasExcluidas": 1,
        "estruturasCompactadas": 1,
        "estruturasOrfas": 0,
        "linhasEstruturas": 2,
    }


def test_estrutura_orfa_exige_confirmacao():
    cadastro, estruturas = cad(USR_A), est(viajante("S3", C, A))
    with pytest.raises(InativacaoError) as erro:
        _executar(cadastro, estruturas, [A])
    assert erro.value.code == "ORFAS_SEM_CONFIRMACAO"
    assert erro.value.status == 400
    assert erro.value.extra["estruturasOrfas"] == ["S3"]

    r = _executar(cadastro, estruturas, [A], ignore_orphan_warning=True)
    s3 = r.estruturas.iloc[0]
    assert (s3["AprovacaoId"], s3["Operacao"], s3["LoginAprovador_1"]) == ("S3", "UPDATE", "")


def test_impressao_digital_divergente_e_recusada():
    with pytest.raises(InativacaoError) as erro:
        InactivationCascadeService.executar(cad(USR_A), est(viajante("S1", A, B)), [A], "0" * 64)
    assert erro.value.code == "ANALISE_DIVERGENTE"
    assert erro.value.status == 409


def test_nenhum_executavel_e_recusado_antes_de_comparar_a_impressao():
    with pytest.raises(InativacaoError) as erro:
        InactivationCascadeService.executar(cad(USR_A), est(viajante("S1", C, B)), [D], "qualquer")
    assert erro.value.code == "NADA_A_EXECUTAR"
    with pytest.raises(InativacaoError) as erro:
        InactivationCascadeService.executar(cad(USR_A), est(viajante("S1", C, B)), [], "qualquer")
    assert erro.value.code == "NADA_A_EXECUTAR"


def test_sem_estruturas_afetadas_o_arquivo_de_estruturas_sai_so_com_cabecalho():
    r = _executar(cad(USR_A), est(viajante("S1", C, B)), [A])
    assert len(r.ficha) == 1
    assert r.estruturas.empty
    assert list(r.estruturas.columns[:2]) == ["Operacao", "AprovacaoId"]


def test_execucao_repetida_da_a_mesma_saida():
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))
    primeira, segunda = _executar(cadastro, estruturas, [A]), _executar(cadastro, estruturas, [A])
    assert primeira.ficha.equals(segunda.ficha)
    assert primeira.estruturas.equals(segunda.estruturas)


def test_nao_altera_as_entradas():
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))
    antes_c, antes_e = cadastro.copy(), estruturas.copy()
    _executar(cadastro, estruturas, [A])
    assert cadastro.equals(antes_c)
    assert estruturas.equals(antes_e)


def test_to_zip_bytes_guarda_cada_arquivo_com_o_nome_dado():
    saida = ExportService.to_zip_bytes({"a.xlsx": io.BytesIO(b"1"), "b.xlsx": io.BytesIO(b"2")})
    with zipfile.ZipFile(saida) as z:
        assert z.namelist() == ["a.xlsx", "b.xlsx"]
        assert z.read("b.xlsx") == b"2"
