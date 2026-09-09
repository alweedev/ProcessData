"""O pipeline de cadastro legado foi aposentado (D4)."""
import importlib

import pandas as pd

from _helpers import valid_cpf


def test_processar_registros_from_files_gone():
    processor = importlib.import_module("backend.processor")
    assert not hasattr(processor, "processar_registros_from_files")
    assert not hasattr(processor, "drop_header_like_rows")
    assert not hasattr(processor, "split_name_first_last")
    assert not hasattr(processor, "sanitize_output_text")


def test_validators_module_gone():
    import pytest

    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("backend.validators")


def test_inativacao_engine_still_works():
    from backend.processor import processar_inativacao_from_paths

    base = pd.DataFrame(
        [{"CPF": valid_cpf(1), "NomeCompleto": "Ana Souza", "Email": "a@x.com", "Status": "ATIVO"}]
    )
    lista = pd.DataFrame([{"CPF": valid_cpf(1)}])
    out, stats = processar_inativacao_from_paths(base, lista)
    assert not out.empty
    assert stats["total_matches"] == 1
