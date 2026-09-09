"""Booleanos exportados sempre como S/N (P6)."""

import pandas as pd
from _helpers import valid_cpf

from backend.processor import processar_inativacao_from_paths
from backend.services.processing_service import ProcessingService


def test_cadastro_sn_mapping(tmp_path):
    variantes = ["Sim", "Não", "YES", "1", ""]
    df = pd.DataFrame(
        [
            {
                "CPF": valid_cpf(i),
                "NOME COMPLETO": f"Pessoa {i}",
                "EMAIL": f"p{i}@x.com",
                "EMPRESA": "Empresa A",
                "Centro de custo": "CC1",
                "SOLICITANTE? (S/N)": v,
            }
            for i, v in enumerate(variantes)
        ]
    )
    p = tmp_path / "cadastro.xlsx"
    df.to_excel(p, index=False)

    _errors, out = ProcessingService.process_records_from_files([str(p)], login_choice="CPF", fluxo="SELF")
    assert out["Solicitante"].tolist() == ["S", "N", "S", "S", "N"]
    assert set(out["Solicitante"].unique()).issubset({"S", "N"})


def test_inativacao_solicitante_not_contaminated_by_master():
    # SolicitanteMaster antes de Solicitante para expor o bug de substring.
    base = pd.DataFrame(
        [
            {
                "CPF": valid_cpf(1),
                "NomeCompleto": "Ana Souza",
                "Status": "ATIVO",
                "SolicitanteMaster": "S",
                "Solicitante": "N",
            }
        ]
    )
    lista = pd.DataFrame([{"CPF": valid_cpf(1)}])
    out, _stats = processar_inativacao_from_paths(base, lista)
    assert out["Solicitante"].tolist() == ["N"]
    assert out["SolicitanteMaster"].tolist() == ["S"]
