from backend.shared.fingerprint import impressao_digital


def test_mesma_informacao_em_outra_ordem_gera_a_mesma_impressao():
    a = {"cpfs": ["2", "1"], "excluidas": ["A"], "compactadas": [{"id": "B", "posicoes": [3, 1]}]}
    b = {"compactadas": [{"posicoes": [1, 3], "id": "B"}], "excluidas": ["A"], "cpfs": ["1", "2"]}
    assert impressao_digital(a) == impressao_digital(b)


def test_impacto_diferente_muda_a_impressao():
    assert impressao_digital({"cpfs": ["1"], "orfas": []}) != impressao_digital({"cpfs": ["1"], "orfas": ["A"]})


def test_aceita_conjuntos_e_devolve_sha256_hex():
    digest = impressao_digital({"cpfs": {"1", "2"}})
    assert len(digest) == 64
    assert int(digest, 16) >= 0
    assert digest == impressao_digital({"cpfs": frozenset({"2", "1"})})
