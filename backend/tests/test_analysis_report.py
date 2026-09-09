import pandas as pd

from backend.services.report_service import ReportService


def test_build_quality_report_counts_rows_and_errors():
    df = pd.DataFrame(
        [
            {
                "Login": "A",
                "NomeCompleto": "ANA",
                "NomeEmpresa": "EMP",
                "CodigoCCustoEmpresa": "1",
                "DescricaoCCustoEmpresa": "CC",
                "Email": "a@x.com",
            },
            {
                "Login": "A",
                "NomeCompleto": "ANA",
                "NomeEmpresa": "",
                "CodigoCCustoEmpresa": "1",
                "DescricaoCCustoEmpresa": "",
                "Email": "",
            },
            {
                "Login": "B",
                "NomeCompleto": "BRUNO",
                "NomeEmpresa": "EMP",
                "CodigoCCustoEmpresa": "2",
                "DescricaoCCustoEmpresa": "CC2",
                "Email": "b@x.com",
            },
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
    assert report["general_errors"] == "Coluna obrigatória ausente: X"


def test_analysis_summary_surfaces_geral_when_required_column_empty(client, tmp_path):
    import pandas as pd
    from _helpers import xlsx_upload

    df = pd.DataFrame([{"NOME COMPLETO": "Ana Um", "EMAIL": "a@x.com"}])
    data = {"files[]": xlsx_upload(df, "in.xlsx"), "login_choice": "EMAIL", "fluxo": "SELF"}
    resp = client.post("/api/analysis/summary", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert resp.get_json()["report"]["general_errors"]
