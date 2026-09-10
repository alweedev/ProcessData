"""InactivationService.search_matches: nomes encontrados não podem cair em not_found."""

import pandas as pd

from backend.services.inactivation_service import InactivationService


def test_found_name_not_listed_in_not_found():
    df_base = pd.DataFrame(
        [
            {"CPF": "", "NomeCompleto": "Maria Silva", "Email": "maria@ex.com", "Status": "ATIVO"},
        ]
    )
    result = InactivationService.search_matches(df_base, ["Maria Silva"])

    assert result["items"][0]["found"] is True
    assert "Maria Silva" not in result["not_found"]
    assert result["not_found"] == []


def test_normalize_lista_columns_detects_accented_header():
    df_lista = pd.DataFrame([{"NÓME COMPLETO": "Maria Silva"}])
    result = InactivationService.normalize_lista_columns(df_lista)
    assert result["NomeCompleto"].tolist() == ["Maria Silva"]


def test_name_not_in_base_is_listed_in_not_found():
    df_base = pd.DataFrame(
        [
            {"CPF": "", "NomeCompleto": "Maria Silva", "Email": "", "Status": "ATIVO"},
        ]
    )
    result = InactivationService.search_matches(df_base, ["Joao Pereira"])

    assert result["items"][0]["found"] is False
    assert result["not_found"] == ["Joao Pereira"]
