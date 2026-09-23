"""Geração de inativação: função e endpoint.

Verificações portadas de backend/test_inativacao_run.py e
backend/test_integration_api.py (scripts manuais removidos em D5).
"""

import pandas as pd
from _helpers import valid_cpf

from backend.processor import processar_inativacao_from_paths


def _base():
    return pd.DataFrame(
        [
            {"CPF": valid_cpf(1), "NomeCompleto": "Maria Silva", "Email": "maria@x.com", "Status": "ATIVO"},
            {"CPF": valid_cpf(2), "NomeCompleto": "Joao Pereira", "Email": "joao@x.com", "Status": "ATIVO"},
            {"CPF": valid_cpf(3), "NomeCompleto": "Mariana Costa", "Email": "mc@x.com", "Status": "INATIVO"},
        ]
    )


def test_match_by_exact_cpf():
    out, stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"CPF": valid_cpf(1)}]))
    assert stats["cpf_matches"] == 1
    assert out["NomeCompleto"].tolist() == ["MARIA SILVA"]
    assert out["Operacao"].tolist() == ["DELETE"]


def test_match_by_exact_full_name():
    out, stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"NomeCompleto": "Maria Silva"}]))
    assert stats["name_matches"] == 1
    assert out["NomeCompleto"].tolist() == ["MARIA SILVA"]


def test_generic_single_name_does_not_over_match():
    out, _stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"NomeCompleto": "Maria"}]))
    # "Maria" != "Maria Silva" nem "Mariana Costa" -> nenhum match exato
    assert out.empty


def test_inactive_row_is_not_matched():
    out, stats = processar_inativacao_from_paths(_base(), pd.DataFrame([{"CPF": valid_cpf(3)}]))
    assert out.empty
    assert stats["cpf_matches"] == 0
