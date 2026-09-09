"""backend.utils foi consolidado em backend.shared.* (D6)."""

import importlib

import pytest


def test_backend_utils_module_removed():
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("backend.utils")


def test_shared_exposes_all_names():
    from backend.shared import cpf_utils, file_utils, text_utils

    assert callable(text_utils.upper_no_accents)
    assert callable(text_utils.normalize_text)
    assert callable(cpf_utils.clean_cpf)
    assert cpf_utils.limpar_cpf_raw is cpf_utils.clean_cpf
    assert callable(cpf_utils.format_cpf_for_output)
    assert callable(file_utils.validar_extensao_arquivo)
    assert callable(file_utils.gerar_nome_arquivo_temporario)


def test_cpf_format_matches_legacy_shape():
    from backend.shared.cpf_utils import format_cpf_for_output, limpar_cpf_raw

    assert limpar_cpf_raw("123.456.789-09") == "12345678909"
    assert format_cpf_for_output("12345678909") == "123456789-09"


def test_no_source_still_imports_backend_utils():
    import pathlib

    root = pathlib.Path(__file__).resolve().parents[2] / "backend"
    offenders = []
    for p in root.rglob("*.py"):
        if p.parent.name == "tests":
            continue
        text = p.read_text(encoding="utf-8")
        if "from backend.utils" in text or "import backend.utils" in text or "from .utils import" in text:
            offenders.append(str(p))
    assert offenders == []
