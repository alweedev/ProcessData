"""Remoção dos parâmetros mortos use_fuzzy/fuzzy_cutoff (C3)."""

import inspect

from backend.processor import processar_inativacao_from_paths
from backend.services.inactivation_service import InactivationService


def test_signatures_have_no_fuzzy_params():
    for fn in (processar_inativacao_from_paths, InactivationService.process_from_dataframes):
        params = set(inspect.signature(fn).parameters)
        assert "use_fuzzy" not in params
        assert "fuzzy_cutoff" not in params
    assert set(inspect.signature(InactivationService.process_from_dataframes).parameters) == {
        "df_base",
        "df_lista",
    }
