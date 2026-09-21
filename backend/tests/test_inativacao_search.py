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


def test_match_by_email_case_insensitive():
    df_base = pd.DataFrame([{"CPF": "", "NomeCompleto": "Maria Silva", "Email": "Maria@Ex.com", "Status": "ATIVO"}])
    result = InactivationService.search_matches(df_base, ["maria@ex.com"])

    assert result["items"][0]["found"] is True
    assert result["items"][0]["email"] == "Maria@Ex.com"


def test_duplicate_cpf_in_input_is_reported_once():
    df_base = pd.DataFrame([{"CPF": "12345678900", "NomeCompleto": "Ana", "Email": "", "Status": "ATIVO"}])
    result = InactivationService.search_matches(df_base, ["12345678900", "123.456.789-00"])

    assert result["duplicates"] == ["12345678900"]
    # o CPF duplicado só deve gerar 1 resultado "found", não 2
    assert sum(1 for item in result["items"] if item["found"]) == 1


def test_busca_por_cpf_restaura_zero_a_esquerda_da_base():
    import pandas as pd

    from backend.services.inactivation_service import InactivationService

    base = pd.DataFrame([{"CPF": "1234567890", "NomeCompleto": "Ana Souza", "Email": "a@x.com", "Status": "ATIVO"}])
    achados = InactivationService.search_matches(base, ["01234567890"])
    assert [i["found"] for i in achados["items"]] == [True]
    assert achados["items"][0]["cpf"] == "01234567890"
