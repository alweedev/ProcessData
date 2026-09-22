import io
import zipfile

import pytest
from _cascade_fixtures import USR_A, A, B, C, D, cad, est, viajante
from openpyxl import load_workbook

from backend.services.export_service import ExportService
from backend.services.inactivation_cascade_service import InactivationCascadeService, InativacaoError
from backend.services.inactivation_service import InactivationService


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


def test_usuario_com_status_em_branco_nao_e_inativado_nem_tem_estruturas_alteradas():
    usr_b_sem_status = (B, "Bruno Lima", "bruno@x.com", "")
    cadastro = cad(USR_A, usr_b_sem_status)
    estruturas = est(viajante("S1", C, A, D), viajante("S2", B, D), viajante("S3", C, B, D))
    r = _executar(cadastro, estruturas, [A, B])
    assert len(r.ficha) == 1
    assert r.resumo["usuariosInativados"] == 1
    assert set(r.estruturas["AprovacaoId"]) == {"S1"}
    with pytest.raises(InativacaoError) as erro:
        _executar(cadastro, estruturas, [B])
    assert erro.value.code == "NADA_A_EXECUTAR"


def test_cpf_repetido_no_cadastro_conta_um_usuario_e_a_ficha_guarda_as_duas_linhas():
    cadastro = cad(USR_A, (A, "Ana Souza", "ana2@x.com", "ATIVO"))
    r = _executar(cadastro, est(viajante("S1", A, B)), [A])
    assert len(r.ficha) == 2
    assert r.resumo["usuariosInativados"] == 1


def test_ficha_que_nao_cobre_todos_os_usuarios_e_erro_interno(monkeypatch):
    cadastro = cad(USR_A, (B, "Bruno Lima", "bruno@x.com", "ATIVO"))
    estruturas = est(viajante("S1", A, C), viajante("S2", B, C))
    real = InactivationService.process_from_dataframes

    def ficha_incompleta(df_base, df_lista):
        ficha, stats = real(df_base, df_lista)
        return ficha.iloc[:1], stats

    monkeypatch.setattr(InactivationService, "process_from_dataframes", staticmethod(ficha_incompleta))
    with pytest.raises(InativacaoError) as erro:
        _executar(cadastro, estruturas, [A, B])
    assert erro.value.code == "ERRO_INTERNO"
    assert erro.value.status == 500


def test_execucao_traz_os_cpfs_executados():
    cadastro = cad(USR_A, (B, "Bruno Lima", "bruno@x.com", "INATIVO"))
    r = _executar(cadastro, est(viajante("S1", A, C)), [A])
    assert r.cpfs == frozenset({A})


def test_estruturas_vazias_gera_planilha_que_abre_no_openpyxl():
    r = _executar(cad(USR_A), est(viajante("S1", C, B)), [A])
    assert r.estruturas.empty
    wb = load_workbook(ExportService.to_excel_bytes(r.estruturas, sheet_name="Aprovacao"))
    cabecalho = [c.value for c in wb["Aprovacao"][1]]
    assert cabecalho[:2] == ["Operacao", "AprovacaoId"]


def test_executar_nunca_inclui_linha_de_estrutura_sem_operacao():
    df_cad = cad(USR_A)
    df_est = est(viajante("X1", A, B))  # A é o viajante, B aprovador em outra estrutura não entra aqui
    analise = InactivationCascadeService.analisar(df_cad, df_est, [A])
    execucao = InactivationCascadeService.executar(df_cad, df_est, [A], analise.payload["impressaoDigital"])
    assert (execucao.estruturas["Operacao"].astype(str).str.strip() != "").all()
