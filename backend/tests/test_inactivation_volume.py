import time

import pandas as pd
from _helpers import valid_cpf

from backend.services.inactivation_cascade_service import InactivationCascadeService

LINHAS = 50_000
USUARIOS = 100
TETO_SEGUNDOS = 8


def _bases():
    inativar = [valid_cpf(i) for i in range(1, USUARIOS + 1)]
    outros = [valid_cpf(i) for i in range(1000, 1030)]
    linhas = [
        {
            "AprovacaoId": f"S{i}",
            "AprovacaoPor": "VIAJANTE" if i % 10 == 0 else "CCEMPRESA",
            "CPF": inativar[(i // 10) % USUARIOS] if i % 10 == 0 else "",
            "LoginAprovador_1": outros[i % 30],
            "LoginAprovador_2": inativar[i % USUARIOS] if i % 3 == 0 else "",
            "LoginAprovador_3": outros[(i + 1) % 30],
        }
        for i in range(LINHAS)
    ]
    cadastro = pd.DataFrame(
        [
            {"CPF": cpf, "NomeCompleto": f"Usuario {n} Teste", "Email": f"u{n}@x.com", "Status": "ATIVO"}
            for n, cpf in enumerate(inativar)
        ]
    )
    return cadastro, pd.DataFrame(linhas).astype(str), inativar


def test_analise_e_execucao_de_50_mil_linhas_com_100_usuarios_cabem_no_teto():
    cadastro, estruturas, cpfs = _bases()

    inicio = time.perf_counter()
    analise = InactivationCascadeService.analisar(cadastro, estruturas, cpfs)
    duracao_analise = time.perf_counter() - inicio
    assert analise.payload["resumo"]["executaveis"] == USUARIOS
    assert duracao_analise < TETO_SEGUNDOS

    inicio = time.perf_counter()
    execucao = InactivationCascadeService.executar(
        cadastro, estruturas, cpfs, analise.payload["impressaoDigital"], ignore_orphan_warning=True
    )
    duracao_execucao = time.perf_counter() - inicio
    assert execucao.resumo["usuariosInativados"] == USUARIOS
    assert duracao_execucao < TETO_SEGUNDOS
