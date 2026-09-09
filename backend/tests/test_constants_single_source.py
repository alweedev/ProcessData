"""MODEL_COLS / FICHA_MAP têm fonte única em backend.domain.rules (D1)."""
from backend import processor
from backend.domain import rules


def test_model_cols_is_the_same_object():
    assert processor.MODEL_COLS is rules.MODEL_COLS
    assert len(rules.MODEL_COLS) == 32


def test_ficha_map_is_the_same_object():
    assert processor.FICHA_MAP is rules.FICHA_MAP
    # a versão canônica mapeia cabeçalhos que a antiga do processor não tinha
    assert rules.FICHA_MAP.get("EMAIL") == "Email"
    assert rules.FICHA_MAP.get("SOLICITANTE") == "Solicitante"


def test_no_local_redefinition_in_processor():
    src = (processor.__file__)
    with open(src, encoding="utf-8") as fh:
        text = fh.read()
    assert "\nMODEL_COLS = [" not in text
    assert "\nFICHA_MAP = {" not in text
