"""Campos obrigatórios da ficha de cadastro, flags S/N e números do relatório de validação."""

import pandas as pd
import pytest
from _helpers import valid_cpf, xlsx_upload

from backend.domain.rules import MODEL_COLS, REQUIRED_FICHA_FIELDS
from backend.services.processing_service import ProcessingService
from backend.services.report_service import ReportService
from backend.services.validation_service import ValidationService


def _row(i=1, **override):
    """Uma linha de ficha completa (todos os obrigatórios preenchidos)."""
    row = {
        "CPF": valid_cpf(i),
        "NOME COMPLETO": f"Pessoa Numero{i}",
        "E-MAIL": f"p{i}@x.com",
        "EMPRESA": "Empresa A",
        "Centro de custo": "CC1",
        "Descrição Centro de Custo": "ADMINISTRATIVO",
        "TELEFONE": "11999990000",
        "Data de Nascimento": "12/05/1990",
        "SOLICITANTE? (S/N)": "S",
    }
    row.update(override)
    return row


def _run(rows, tmp_path, login_choice="CPF", fluxo="SELF", name="in.xlsx"):
    path = tmp_path / name
    pd.DataFrame(rows).to_excel(path, index=False)
    return ProcessingService.process_records_from_files([str(path)], login_choice=login_choice, fluxo=fluxo)


def _report(rows, tmp_path, **kwargs):
    errors, df = _run(rows, tmp_path, **kwargs)
    return ReportService.build_quality_report(df, errors), errors, df


# ------------------------------------------------------------ obrigatórios


def test_ficha_completa_nao_tem_erro_algum(tmp_path):
    report, errors, _df = _report([_row(1), _row(2)], tmp_path)

    assert errors == {}
    assert report["valid_rows"] == 2 and report["invalid_rows"] == 0
    assert not any(report["required_blank"].values())


def test_os_oito_obrigatorios_sao_os_definidos_pelo_negocio():
    assert list(REQUIRED_FICHA_FIELDS.values()) == [
        "CPF",
        "Empresa",
        "Centro de custo - Código",
        "Centro de custo - Descrição",
        "Nome completo",
        "E-mail",
        "Telefone",
        "Data de nascimento",
    ]


@pytest.mark.parametrize(
    ("coluna", "rotulo"),
    [
        ("CPF", "CPF"),
        ("EMPRESA", "Empresa"),
        ("Centro de custo", "Centro de custo - Código"),
        ("Descrição Centro de Custo", "Centro de custo - Descrição"),
        ("NOME COMPLETO", "Nome completo"),
        ("E-MAIL", "E-mail"),
        ("TELEFONE", "Telefone"),
        ("Data de Nascimento", "Data de nascimento"),
    ],
)
def test_obrigatorio_em_branco_invalida_a_linha_e_aparece_no_relatorio(tmp_path, coluna, rotulo):
    # a 2ª linha tem o campo em branco; a 1ª segue completa (a coluna não fica "toda vazia")
    report, errors, _df = _report([_row(1), _row(2, **{coluna: ""})], tmp_path)

    assert report["invalid_rows"] == 1 and report["valid_rows"] == 1
    assert errors[1] == f"Campo obrigatório em branco: {rotulo}"
    assert report["required_blank"][rotulo] == 1


def test_campos_opcionais_em_branco_nao_invalidam(tmp_path):
    rows = [
        _row(
            1, **{"MATRICULA": "", "CARGO": "", "DEPARTAMENTO": "", "NIVEL": "", "SOBRENOME (limite 50 caracteres)": ""}
        )
    ]

    _report_, errors, _df = _report(rows, tmp_path)

    assert errors == {}


def test_cpf_com_poucos_digitos_continua_invalido(tmp_path):
    report, errors, _df = _report([_row(1), _row(2, CPF="123")], tmp_path)

    assert errors[1] == "CPF deve ter 11 dígitos"
    assert report["invalid_rows"] == 1


def test_email_preenchido_precisa_ter_formato_valido(tmp_path):
    _report_, errors, _df = _report([_row(1, **{"E-MAIL": "fulano@empresa"})], tmp_path)

    assert errors[0] == "Email inválido"


def test_cpf_e_exigido_mesmo_quando_o_login_e_por_email(tmp_path):
    _report_, errors, _df = _report([_row(1, CPF="")], tmp_path, login_choice="EMAIL")

    assert errors[0] == "Campo obrigatório em branco: CPF"


# --------------------------------------------- colunas ausentes / vazias (geral)


def test_coluna_obrigatoria_ausente_na_ficha_e_apontada_pelo_nome(tmp_path):
    rows = [{k: v for k, v in _row(i).items() if k not in ("TELEFONE", "Data de Nascimento")} for i in (1, 2)]

    report, errors, _df = _report(rows, tmp_path)

    assert "Coluna obrigatória ausente na ficha: Telefone" in report["general_errors"]
    assert "Coluna obrigatória ausente na ficha: Data de nascimento" in report["general_errors"]
    assert report["invalid_rows"] == 2  # e cada linha também acusa o que falta
    assert errors[0] == "Campo obrigatório em branco: Telefone; Campo obrigatório em branco: Data de nascimento"


def test_coluna_obrigatoria_toda_vazia_e_apontada_como_vazia(tmp_path):
    report, _errors, _df = _report([_row(1, TELEFONE=""), _row(2, TELEFONE="")], tmp_path)

    assert report["general_errors"] == "Coluna obrigatória vazia na ficha: Telefone"


def test_validate_dataframe_nao_repete_ausente_como_vazia():
    df = pd.DataFrame([{"Telefone": ""}])
    msgs = ValidationService.validate_dataframe(df, ausentes=["Telefone"])

    assert msgs == ["Coluna obrigatória ausente na ficha: Telefone"]


# --------------------------------------------------------- Data de nascimento


@pytest.mark.parametrize(
    "cabecalho", ["Data de Nascimento", "DATA NASCIMENTO", "data de nascimento", "Dt. Nascimento", "NASCIMENTO"]
)
def test_data_de_nascimento_e_lida_com_varios_cabecalhos(tmp_path, cabecalho):
    row = _row(1)
    del row["Data de Nascimento"]
    row[cabecalho] = "12/05/1990"

    report, errors, _df = _report([row], tmp_path)

    assert errors == {}
    assert report["required_blank"]["Data de nascimento"] == 0


def test_data_de_nascimento_valida_mas_nao_vai_para_o_arquivo_de_carga(tmp_path):
    _errors, df = _run([_row(1)], tmp_path)

    assert list(df.columns) == MODEL_COLS
    assert "DataNascimento" not in df.columns and "CPF" not in df.columns


# ------------------------------------------- Telefone: "CELULAR - CONTATO" de reserva


def test_telefone_em_branco_usa_o_celular_contato_e_ele_vai_para_o_arquivo(tmp_path):
    row = _row(1, TELEFONE="")
    row["CELULAR - CONTATO"] = "11988887777"

    report, errors, df = _report([row], tmp_path)

    assert errors == {}
    assert df["Telefone"].tolist() == ["11988887777"]
    assert report["required_blank"]["Telefone"] == 0


def test_telefone_preenchido_tem_prioridade_sobre_o_celular_contato(tmp_path):
    row = _row(1, TELEFONE="1133334444")
    row["CELULAR - CONTATO"] = "11988887777"

    _errors, df = _run([row], tmp_path)

    assert df["Telefone"].tolist() == ["1133334444"]


def test_so_as_linhas_com_telefone_em_branco_recebem_o_celular_contato(tmp_path):
    rows = [_row(1, TELEFONE="1133334444"), _row(2, TELEFONE=""), _row(3, TELEFONE="")]
    for row, celular in zip(rows, ["", "11977776666", ""], strict=True):
        row["CELULAR - CONTATO"] = celular

    report, errors, df = _report(rows, tmp_path)

    assert df["Telefone"].tolist() == ["1133334444", "11977776666", ""]
    assert errors == {2: "Campo obrigatório em branco: Telefone"}  # sem telefone nem contato: continua inválida
    assert report["required_blank"]["Telefone"] == 1


def test_sem_a_coluna_telefone_mas_com_celular_contato_nao_ha_coluna_ausente(tmp_path):
    rows = []
    for i in (1, 2):
        row = {k: v for k, v in _row(i).items() if k != "TELEFONE"}
        row["CELULAR - CONTATO"] = f"1198888000{i}"
        rows.append(row)

    report, errors, df = _report(rows, tmp_path)

    assert errors == {} and report["general_errors"] == ""
    assert df["Telefone"].tolist() == ["11988880001", "11988880002"]


def test_sem_telefone_e_com_celular_contato_todo_vazio_a_coluna_e_apontada_como_vazia(tmp_path):
    rows = []
    for i in (1, 2):
        row = {k: v for k, v in _row(i).items() if k != "TELEFONE"}
        row["CELULAR - CONTATO"] = ""
        rows.append(row)

    report, _errors, _df = _report(rows, tmp_path)

    assert report["general_errors"] == "Coluna obrigatória vazia na ficha: Telefone"


@pytest.mark.parametrize("cabecalho", ["CELULAR - CONTATO", "Celular - Contato", "celular contato", "CELULAR-CONTATO"])
def test_celular_contato_e_lido_com_varios_cabecalhos(tmp_path, cabecalho):
    row = _row(1, TELEFONE="")
    row[cabecalho] = "11988887777"

    _errors, df = _run([row], tmp_path)

    assert df["Telefone"].tolist() == ["11988887777"]


def test_celular_contato_nao_vira_coluna_do_arquivo_de_carga(tmp_path):
    row = _row(1, TELEFONE="")
    row["CELULAR - CONTATO"] = "11988887777"

    _errors, df = _run([row], tmp_path)

    assert list(df.columns) == MODEL_COLS
    assert "CelularContato" not in df.columns


# ------------------------------------------------------------------ flags S/N


@pytest.mark.parametrize(
    ("valor", "esperado"),
    [("S", "S"), ("Sim", "S"), ("sim", "S"), ("N", "N"), ("Não", "N"), ("", "N"), ("talvez", "N")],
)
def test_solicitante_sim_vira_s_nao_vira_n_e_branco_vai_para_n_sem_invalidar(tmp_path, valor, esperado):
    errors, df = _run([_row(1, **{"SOLICITANTE? (S/N)": valor})], tmp_path)

    assert errors == {}
    assert df["Solicitante"].tolist() == [esperado]


def test_sem_a_coluna_solicitante_todas_as_linhas_saem_com_n_e_validas(tmp_path):
    rows = [{k: v for k, v in _row(i).items() if k != "SOLICITANTE? (S/N)"} for i in (1, 2)]

    errors, df = _run(rows, tmp_path)

    assert errors == {}
    assert df["Solicitante"].tolist() == ["N", "N"]


def test_terceiro_em_branco_vai_para_n(tmp_path):
    errors, df = _run([_row(1, **{"TERCEIRO? (S/N)": ""}), _row(2, **{"TERCEIRO? (S/N)": "Sim"})], tmp_path)

    assert errors == {}
    assert df["Terceiro"].tolist() == ["N", "S"]


# ------------------------------------------------------- números do relatório


def test_linhas_repetidas_sao_contadas_como_duplicadas_e_removidas(tmp_path):
    report, _errors, df = _report([_row(1), _row(1), _row(2)], tmp_path)

    assert report["duplicated_rows"] == 1
    assert report["total_rows"] == 2 == len(df)  # a repetida sai do arquivo, o total é o que sai
    assert report["valid_rows"] == 2


def test_linhas_em_branco_nao_sao_duplicadas_nem_invalidas(tmp_path):
    vazia = {k: "" for k in _row(1)}

    report, errors, _df = _report([_row(1), vazia, vazia, _row(3)], tmp_path)

    assert report["duplicated_rows"] == 0
    assert report["invalid_rows"] == 0 and errors == {}
    assert report["total_rows"] == 2 and report["valid_rows"] == 2


def test_erro_de_linha_repetida_descartada_nao_conta_como_invalida(tmp_path):
    report, _errors, _df = _report([_row(1, TELEFONE=""), _row(1, TELEFONE=""), _row(2)], tmp_path)

    assert report["duplicated_rows"] == 1
    assert report["total_rows"] == 2
    assert report["invalid_rows"] == 1  # só a que ficou; a repetida saiu junto com o erro


def test_erros_apontam_a_linha_do_excel_nao_o_indice_interno(tmp_path):
    # cabeçalho = linha 1; a linha com problema é a 3ª de dados = linha 4 do Excel
    report, _errors, _df = _report([_row(1), _row(2), _row(3, TELEFONE=""), _row(4)], tmp_path)

    assert list(report["line_errors"]) == ["Linha 4"]


def test_com_varios_arquivos_o_erro_diz_qual_arquivo(tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    pd.DataFrame([_row(1), _row(2)]).to_excel(a, index=False)
    pd.DataFrame([_row(3), _row(4, TELEFONE="")]).to_excel(b, index=False)

    errors, df = ProcessingService.process_records_from_files([str(a), str(b)])
    report = ReportService.build_quality_report(df, errors)

    assert list(report["line_errors"]) == ["Arquivo 2 · linha 3"]


def test_relatorio_sem_attrs_do_pipeline_continua_funcionando():
    # DataFrame montado à mão (sem attrs): usa só o que há nas colunas
    df = pd.DataFrame(
        [{"Login": "A", "NomeCompleto": "ANA", "Telefone": ""}, {"Login": "A", "NomeCompleto": "ANA", "Telefone": "1"}]
    )

    report = ReportService.build_quality_report(df, {0: "x"})

    assert report["duplicated_rows"] == 1
    assert report["required_blank"] == {"Nome completo": 0, "Telefone": 1}
    assert list(report["line_errors"]) == ["0"]


# ----------------------------------------------------------------- pela API


def test_api_de_validacao_devolve_relatorio_consistente(client):
    df = pd.DataFrame([_row(1), _row(1), _row(2, TELEFONE=""), {k: "" for k in _row(1)}])
    data = {"files[]": xlsx_upload(df, "fichas.xlsx"), "login_choice": "CPF", "fluxo": "SELF"}

    resp = client.post("/api/analysis/summary", data=data, content_type="multipart/form-data")

    assert resp.status_code == 200, resp.get_data(as_text=True)
    report = resp.get_json()["report"]
    assert report["total_rows"] == 2  # a repetida e a em branco saem
    assert report["duplicated_rows"] == 1
    assert report["invalid_rows"] == 1 and report["valid_rows"] == 1
    assert report["line_errors"] == {"Linha 4": "Campo obrigatório em branco: Telefone"}
    assert report["required_blank"]["Telefone"] == 1


def test_geracao_nao_e_bloqueada_por_linhas_invalidas(client):
    # a validação é informativa: quem decide gerar é o usuário
    df = pd.DataFrame([_row(1), _row(2, TELEFONE="")])
    data = {"files[]": xlsx_upload(df, "fichas.xlsx"), "login_choice": "CPF", "fluxo": "SELF"}

    resp = client.post("/api/process_cadastro", data=data, content_type="multipart/form-data")

    assert resp.status_code == 200


# ------------------------------------------- problemas por passageiro (tela do resultado)


def test_relatorio_traz_cada_problema_separado_com_o_nome_do_passageiro(tmp_path):
    rows = [_row(1), _row(2, TELEFONE="", **{"E-MAIL": "sem-arroba"}), _row(3, CPF="123")]

    report, _errors, _df = _report(rows, tmp_path)

    assert report["line_details"] == [
        {
            "label": "Linha 3",
            "nome": "PESSOA NUMERO2",
            "erros": ["Campo obrigatório em branco: Telefone", "Email inválido"],
            "sem_preenchimento": False,
        },
        {
            "label": "Linha 4",
            "nome": "PESSOA NUMERO3",
            "erros": ["CPF deve ter 11 dígitos"],
            "sem_preenchimento": False,
        },
    ]
    assert report["line_errors"]["Linha 3"] == "Campo obrigatório em branco: Telefone; Email inválido"  # segue igual


def test_linha_com_problema_e_sem_nome_vem_com_o_nome_vazio(tmp_path):
    report, _errors, _df = _report([_row(1), _row(2, **{"NOME COMPLETO": ""})], tmp_path)

    [detalhe] = report["line_details"]
    assert detalhe["nome"] == "" and detalhe["erros"] == ["Campo obrigatório em branco: Nome completo"]


def test_planilha_sem_problemas_nao_tem_detalhes(tmp_path):
    report, _errors, _df = _report([_row(1), _row(2)], tmp_path)

    assert report["line_details"] == []


def test_relatorio_montado_sem_attrs_do_pipeline_tambem_traz_detalhes():
    df = pd.DataFrame([{"Login": "A", "NomeCompleto": "ANA"}])

    report = ReportService.build_quality_report(df, {0: "x; y"})

    assert report["line_details"] == [{"label": "0", "nome": "", "erros": ["x", "y"], "sem_preenchimento": False}]


def test_linha_so_com_o_nome_e_marcada_como_sem_preenchimento(tmp_path):
    so_nome = {k: "" for k in _row(1)} | {"NOME COMPLETO": "Só Nome"}
    rows = [_row(1), so_nome, _row(3, TELEFONE=""), _row(4, TELEFONE="", CPF="", EMPRESA="")]

    report, _errors, _df = _report(rows, tmp_path)

    marcas = {d["label"]: d["sem_preenchimento"] for d in report["line_details"]}
    assert marcas == {"Linha 3": True, "Linha 4": False, "Linha 5": False}  # faltar alguns não é "sem preenchimento"


def test_estrangeiro_com_passaporte_e_sem_cpf_nao_e_linha_sem_preenchimento(tmp_path):
    row = {k: "" for k in _row(1)} | {"NOME COMPLETO": "Outro Nome", "NÚMERO PASSAPORTE": "AB123"}

    report, _errors, _df = _report([_row(1), row], tmp_path, login_choice="EMAIL")

    [detalhe] = report["line_details"]
    assert "Campo obrigatório em branco: CPF" not in detalhe["erros"]  # o passaporte dispensa o CPF
    assert detalhe["sem_preenchimento"] is False
