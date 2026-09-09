"""Sanidade do harness de testes (fixtures, isolamento de estado)."""

import os

from _helpers import valid_cpf

from backend.core.config import settings


def test_health_ok(client):
    resp = client.get("/api/health")
    assert resp.status_code == 200
    assert resp.get_json()["status"] == "OK"


def test_state_is_isolated_under_tmp(tmp_path):
    assert str(tmp_path) in settings.UPLOAD_FOLDER
    assert os.path.isdir(settings.UPLOAD_FOLDER)
    assert str(tmp_path) in settings.HISTORY_LOG_FILE


def test_valid_cpf_helper_has_correct_check_digits():
    cpf = valid_cpf(3)
    assert len(cpf) == 11 and cpf.isdigit()
    # revalida com o mesmo algoritmo de dígito verificador
    digits = [int(c) for c in cpf]
    dv1 = sum(d * w for d, w in zip(digits[:9], range(10, 1, -1))) % 11
    dv1 = 0 if dv1 < 2 else 11 - dv1
    dv2 = sum(d * w for d, w in zip(digits[:10], range(11, 1, -1))) % 11
    dv2 = 0 if dv2 < 2 else 11 - dv2
    assert digits[9] == dv1 and digits[10] == dv2
