"""Comportamentos legados portados para ProcessingService (D2 + D3)."""
import pandas as pd

from backend.services.processing_service import ProcessingService

from _helpers import valid_cpf


def _run(rows, tmp_path, login_choice="CPF", fluxo="SELF"):
    df = pd.DataFrame(rows)
    p = tmp_path / "in.xlsx"
    df.to_excel(p, index=False)
    return ProcessingService.process_records_from_files([str(p)], login_choice=login_choice, fluxo=fluxo)


# ---------------------------------------------------------------- D2

def test_terceiro_digits_preserved(tmp_path):
    rows = [
        {"CPF": valid_cpf(1), "NOME COMPLETO": "Ana Um", "EMAIL": "a@x.com", "SOLICITANTE? (S/N)": "S", "Terceiro": "123"},
        {"CPF": valid_cpf(2), "NOME COMPLETO": "Bruno Dois", "EMAIL": "b@x.com", "SOLICITANTE? (S/N)": "S", "Terceiro": "Sim"},
        {"CPF": valid_cpf(3), "NOME COMPLETO": "Caio Tres", "EMAIL": "c@x.com", "SOLICITANTE? (S/N)": "S", "Terceiro": "Nao"},
    ]
    _err, out = _run(rows, tmp_path)
    assert out["Terceiro"].tolist() == ["123", "S", "N"]


def test_bool_accent_and_fullwidth_folding(tmp_path):
    rows = [
        {"CPF": valid_cpf(1), "NOME COMPLETO": "Ana Um", "EMAIL": "a@x.com", "SOLICITANTE? (S/N)": "ＳＩＭ"},  # SIM fullwidth
        {"CPF": valid_cpf(2), "NOME COMPLETO": "Bruno Dois", "EMAIL": "b@x.com", "SOLICITANTE? (S/N)": "Não"},
    ]
    _err, out = _run(rows, tmp_path)
    assert out["Solicitante"].tolist() == ["S", "N"]


def test_front_prefix_no_nan_or_bare_prefix(tmp_path):
    rows = [
        {"NOME COMPLETO": "Sem Documento", "EMAIL": "", "EMPRESA": "E", "SOLICITANTE? (S/N)": "S"},
    ]
    _err, out = _run(rows, tmp_path, login_choice="CPF", fluxo="FRONT")
    logins = out["Login"].tolist()
    assert all("nan" not in str(v).lower() for v in logins)
    assert "FRONT" not in logins  # não vira só o prefixo
