"""MODEL_COLS / FICHA_MAP têm fonte única em backend.domain.rules (D1)."""

from backend import processor
from backend.domain import rules
from backend.services.processing_service import FICHA_MAP as SERVICE_FICHA_MAP


def test_model_cols_is_the_same_object():
    assert processor.MODEL_COLS is rules.MODEL_COLS
    assert len(rules.MODEL_COLS) == 32


def test_ficha_map_single_source():
    # cadastro (services) usa exatamente o FICHA_MAP de domain.rules
    assert SERVICE_FICHA_MAP is rules.FICHA_MAP
    # e o processor não redefine nem reexporta o seu próprio
    assert not hasattr(processor, "FICHA_MAP")
    assert rules.FICHA_MAP.get("EMAIL") == "Email"
    assert rules.FICHA_MAP.get("SOLICITANTE") == "Solicitante"


def test_no_local_redefinition_in_processor():
    with open(processor.__file__, encoding="utf-8") as fh:
        text = fh.read()
    assert "\nMODEL_COLS = [" not in text
    assert "\nFICHA_MAP = {" not in text
