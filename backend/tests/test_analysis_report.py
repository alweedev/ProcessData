import pandas as pd

from backend.services.report_service import ReportService


def test_build_quality_report_counts_rows_and_errors():
    df = pd.DataFrame(
        [
            {"Login": "A", "NomeCompleto": "ANA", "NomeEmpresa": "EMP", "CodigoCCustoEmpresa": "1", "DescricaoCCustoEmpresa": "CC", "Email": "a@x.com"},
            {"Login": "A", "NomeCompleto": "ANA", "NomeEmpresa": "", "CodigoCCustoEmpresa": "1", "DescricaoCCustoEmpresa": "", "Email": ""},
            {"Login": "B", "NomeCompleto": "BRUNO", "NomeEmpresa": "EMP", "CodigoCCustoEmpresa": "2", "DescricaoCCustoEmpresa": "CC2", "Email": "b@x.com"},
        ]
    )

    errors = {
        1: "NomeEmpresa ausente",
        "__geral__": "Coluna obrigatória ausente: X",
    }

    report = ReportService.build_quality_report(df, errors)

    assert report["total_rows"] == 3
    assert report["invalid_rows"] == 1
    assert report["valid_rows"] == 2
    assert report["duplicated_rows"] >= 1
    assert "NomeEmpresa" in report["required_blank"]
