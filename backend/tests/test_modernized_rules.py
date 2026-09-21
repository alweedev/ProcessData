import os
import tempfile

import pandas as pd

from backend.domain.rules import MODEL_COLS
from backend.processor import processar_inativacao_from_paths
from backend.services.processing_service import ProcessingService
from backend.services.validation_service import ValidationService
from backend.shared.cpf_utils import clean_cpf, format_cpf_for_output
from backend.shared.text_utils import sanitize_output_text, split_name_first_last, upper_no_accents


def test_cpf_normalization_and_formatting():
    assert clean_cpf("123.456.789-09") == "12345678909"
    assert format_cpf_for_output("12345678909") == "123456789-09"


def test_clean_cpf_strips_float_stringified_excel_cell():
    # openpyxl/pandas podem devolver célula numérica como "12345678909.0"
    assert clean_cpf("12345678909.0") == "12345678909"


def test_clean_cpf_restores_lost_leading_zero():
    # Excel tratando CPF como número descarta o zero à esquerda
    assert clean_cpf("1234567890") == "01234567890"


def test_text_normalization_and_name_split():
    assert upper_no_accents("José da Silva") == "JOSE DA SILVA"
    assert sanitize_output_text("João / Silva", None) == "JOAO / SILVA"
    first, last = split_name_first_last("Maria Joao Silva")
    assert first == "MARIA"
    assert last == "SILVA"


def test_validation_rejects_missing_name_and_ignores_solicitante():
    row = {
        "Solicitante": "X",  # sim/não nunca invalida: o pipeline normaliza depois
        "NomeCompleto": "",
        "Email": "",
        "Nivel": "OPERACIONAL",
        "CPF": "12345678909",
    }
    errors = ValidationService.validate_row(row)
    assert "Campo obrigatório em branco: Nome completo" in errors
    assert not any("Solicitante" in msg for msg in errors)


def test_validation_accepts_cpf_missing_one_leading_zero():
    # Excel derrubou 1 zero à esquerda (10 dígitos reais) -> zfill restaura;
    # continua sendo um CPF plausível, não deve gerar erro.
    row = {"Solicitante": "S", "NomeCompleto": "Ana Souza", "Email": "", "CPF": "1234567890"}
    errors = ValidationService.validate_row(row)
    assert not any("CPF" in msg for msg in errors)


def test_validation_rejects_cpf_with_too_few_digits():
    # Menos de 10 dígitos não é "1 zero perdido pelo Excel", é entrada
    # incompleta -- não deve virar um CPF "válido" zero-preenchido.
    row = {"Solicitante": "S", "NomeCompleto": "Ana Souza", "Email": "", "CPF": "12345"}
    errors = ValidationService.validate_row(row)
    assert any("CPF deve ter 11 dígitos" in msg for msg in errors)


def test_processing_service_generates_output_for_self_flow():
    df = pd.DataFrame(
        [
            {
                "CPF": "11122233344",
                "NomeCompleto": "Ana Souza",
                "Solicitante": "S",
                "Email": "ana@empresa.com",
                "Empresa": "Empresa A",
                "Centro de custo": "CC1",
                "Descrição Centro de Custo": "ADMINISTRATIVO",
                "Telefone": "11999990001",
                "Data de Nascimento": "12/05/1990",
            },
            {
                "CPF": "22233344455",
                "NomeCompleto": "Bruno Lima",
                "Solicitante": "N",
                "Email": "bruno@empresa.com",
                "Empresa": "Empresa A",
                "Centro de custo": "CC1",
                "Descrição Centro de Custo": "ADMINISTRATIVO",
                "Telefone": "11999990002",
                "Data de Nascimento": "03/11/1985",
            },
        ]
    )

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        df.to_excel(tmp.name, index=False)
        tmp_path = tmp.name

    try:
        errors, result = ProcessingService.process_records_from_files([tmp_path], login_choice="CPF", fluxo="SELF")
        assert not errors
        assert list(result.columns) == MODEL_COLS
        assert set(result["Solicitante"].unique()).issubset({"S", "N"})
        assert result["ViajanteMasterNacional"].tolist() == ["N", "N"]
        assert result["Login"].tolist()[0].startswith("111222333-") or result["Login"].tolist()[0].startswith(
            "111222333"
        )
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def test_nivel_correction_from_validate_row_is_persisted_in_output():
    """Nivel='OPER' deve ser normalizado para 'OPERACIONAL' na planilha final,
    não apenas na cópia interna usada pela validação (row de iterrows())."""
    df = pd.DataFrame(
        [
            {
                "CPF": "11122233344",
                "NOME COMPLETO": "Ana Souza",
                "SOLICITANTE? (S/N)": "S",
                "EMAIL": "ana@empresa.com",
                "EMPRESA": "Empresa A",
                "NIVEL": "OPER",
            },
        ]
    )

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        df.to_excel(tmp.name, index=False)
        tmp_path = tmp.name

    try:
        _errors, result = ProcessingService.process_records_from_files([tmp_path], login_choice="CPF", fluxo="SELF")
        assert result["Nivel"].tolist() == ["OPERACIONAL"]
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def test_inactivation_boolean_fields_are_exported_as_s_or_n():
    base = pd.DataFrame(
        [
            {
                "CPF": "11122233344",
                "NomeCompleto": "Ana Souza",
                "Status": "ATIVO",
                "Solicitante": "Sim",
                "Terceiro": "Não",
                "ViajanteMasterNacional": "Sim",
                "ViajanteMasterInternacional": "Não",
            }
        ]
    )
    lista = pd.DataFrame([{"CPF": "11122233344"}])

    result, _ = processar_inativacao_from_paths(base, lista)

    boolean_columns = [
        "Solicitante",
        "Vip",
        "ViajanteMasterNacional",
        "ViajanteMasterInternacional",
        "SolicitanteMaster",
        "MasterAdiantamento",
        "MasterReembolso",
        "Terceiro",
    ]
    for column in boolean_columns:
        assert set(result[column].unique()).issubset({"S", "N"})

    assert result.loc[0, "Solicitante"] == "S"
    assert result.loc[0, "Terceiro"] == "N"
    assert result.loc[0, "ViajanteMasterNacional"] == "S"
    assert result.loc[0, "ViajanteMasterInternacional"] == "N"
