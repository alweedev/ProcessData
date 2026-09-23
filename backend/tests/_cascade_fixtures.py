"""Builders de planilhas para os testes da inativação em cascata."""

import pandas as pd
from _helpers import valid_cpf

A, B, C, D = (valid_cpf(i) for i in (1, 2, 3, 4))

USR_A = (A, "Ana Souza", "ana@x.com", "ATIVO")
USR_B = (B, "Bruno Lima", "bruno@x.com", "ATIVO")
USR_C = (C, "Carla Dias", "carla@x.com", "ATIVO")


def cad(*usuarios):
    """Base de cadastro; cada usuário é `(cpf, nome, email, status)`."""
    return pd.DataFrame([{"CPF": c, "NomeCompleto": n, "Email": e, "Status": s} for c, n, e, s in usuarios]).astype(str)


def est(*linhas):
    """Base de estruturas a partir de linhas (dicts); células ausentes viram texto vazio."""
    return pd.DataFrame(list(linhas)).fillna("").astype(str)


def viajante(aid, cpf, *aprovadores, segundo=""):
    """Estrutura AprovacaoPor=VIAJANTE do viajante `cpf`, com os aprovadores em 1..n."""
    linha = {"AprovacaoId": aid, "AprovacaoPor": "VIAJANTE", "CPF": cpf, "NomeViajante": f"Viajante {aid}"}
    for i, login in enumerate(aprovadores, 1):
        linha[f"LoginAprovador_{i}"] = login
    if segundo:
        linha["LoginAprovador_SEGUNDO_NIVEL"] = segundo
    return linha
