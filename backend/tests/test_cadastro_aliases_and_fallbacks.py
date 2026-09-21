"""Aliases de cabeçalho, e-mail de login, passaporte, formato de e-mail, planilhas vazias e parâmetros."""

import io
import zipfile

import openpyxl
import pandas as pd
import pytest
from _helpers import valid_cpf, xlsx_upload

from backend.domain.rules import FICHA_MAP, MODEL_COLS
from backend.services.processing_service import ProcessingService
from backend.services.report_service import ReportService
from backend.services.validation_service import ValidationService
from backend.shared import excel_reader
from backend.shared.text_utils import upper_no_accents


def _row(i=1, **override):
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


def _write(rows, path):
    pd.DataFrame(rows).to_excel(path, index=False)
    return str(path)


def _run(rows, tmp_path, login_choice="CPF", fluxo="SELF"):
    path = _write(rows, tmp_path / "in.xlsx")
    errors, df = ProcessingService.process_records_from_files([path], login_choice=login_choice, fluxo=fluxo)
    return errors, df, ReportService.build_quality_report(df, errors)


# ------------------------------------------------------------------ aliases


@pytest.mark.parametrize(
    ("cabecalho", "campo"),
    [
        ("NOME COMPLETO (até 50 caracteres)", "NomeCompleto"),
        ("nome completo (ATE 50 caracteres)", "NomeCompleto"),
        ("NOME (limite 20 caracteres)", "Nome"),
        ("SOBRENOME (limite 20 caracteres)", "SobreNome"),
    ],
)
def test_aliases_portugues_das_fichas_de_clientes(cabecalho, campo):
    assert {upper_no_accents(k).strip(): v for k, v in FICHA_MAP.items()}[upper_no_accents(cabecalho).strip()] == campo


def test_ficha_com_nome_completo_ate_50_caracteres_passa_como_as_demais(tmp_path):
    row = _row(1)
    row["NOME COMPLETO (até 50 caracteres)"] = row.pop("NOME COMPLETO")

    errors, df, _report = _run([row], tmp_path)

    assert errors == {}
    assert df["NomeCompleto"].tolist() == ["PESSOA NUMERO1"]


def test_ficha_em_ingles_e_reconhecida_e_os_valores_saem_nos_campos_certos(tmp_path):
    row = {
        "Company": "Empresa A",
        "Cost center code": "CC9",
        "Cost center description": "FINANCEIRO",
        "First Name": "Ana",
        "Last name (20 caracteres)": "Souza",
        "Full Name (50 caracteres)": "Ana Souza",
        "Email": "ana@x.com",
        "Mobile Number": "11988887777",
        "Position": "Analista",
        "Department": "TI",
        "Applicant? (Y/N)": "Y",
        "Third Part? (Y/N)": "N",
        "Passport": "AB123456",
        "Birth date": "12/05/1990",
    }

    errors, df, report = _run([row], tmp_path, login_choice="EMAIL")

    assert errors == {} and report["general_errors"] == ""
    out = df.iloc[0]
    assert (out["NomeEmpresa"], out["CodigoCCustoEmpresa"], out["DescricaoCCustoEmpresa"]) == (
        "EMPRESA A",
        "CC9",
        "FINANCEIRO",
    )
    assert (out["Nome"], out["SobreNome"], out["NomeCompleto"]) == ("ANA", "SOUZA", "ANA SOUZA")
    assert (out["Telefone"], out["Cargo"], out["Departamento"]) == ("11988887777", "ANALISTA", "TI")
    assert (out["Solicitante"], out["Terceiro"], out["Login"]) == ("S", "N", "ANA@X.COM")


def test_ficha_em_ingles_com_login_por_cpf_continua_pedindo_cpf(tmp_path):
    row = {
        "Company": "E",
        "Cost center code": "1",
        "Cost center description": "D",
        "Full Name (50 caracteres)": "Ana Souza",
        "Email": "a@x.com",
        "Mobile Number": "1",
        "Passport": "AB1",
        "Birth date": "1",
    }

    _errors, _df, report = _run([row], tmp_path, login_choice="CPF")

    assert report["invalid_rows"] == 1
    assert "Campo obrigatório em branco: CPF" in report["line_errors"]["Linha 2"]


# ------------------------------------------------------------- e-mail de login


def test_email_de_login_prevalece_sobre_o_email_normal(tmp_path):
    row = _row(1, **{"E-MAIL": "contato@x.com"})
    row["E-MAIL (LOGIN)"] = "login@x.com"

    errors, df, _report = _run([row], tmp_path, login_choice="EMAIL")

    assert errors == {}
    assert df["Email"].tolist() == ["LOGIN@X.COM"]
    assert df["Login"].tolist() == ["LOGIN@X.COM"]


def test_email_de_login_em_branco_usa_o_email_normal(tmp_path):
    row = _row(1, **{"E-MAIL": "contato@x.com"})
    row["LOGIN (E-MAIL CORPORATIVO)"] = ""

    _errors, df, _report = _run([row], tmp_path, login_choice="EMAIL")

    assert df["Email"].tolist() == ["CONTATO@X.COM"]


def test_ficha_so_com_email_de_login_nao_conta_email_como_coluna_ausente(tmp_path):
    row = _row(1)
    row["E-MAIL (LOGIN)"] = row.pop("E-MAIL")

    errors, df, report = _run([row], tmp_path)

    assert errors == {} and report["general_errors"] == ""
    assert df["Email"].tolist() == ["P1@X.COM"]


# -------------------------------------------------------- CPF x passaporte


def _estrangeiro(**override):
    row = _row(1, CPF="")
    row["NÚMERO PASSAPORTE"] = "AB123456"
    row.update(override)
    return row


def test_passaporte_dispensa_o_cpf_quando_o_login_e_por_email(tmp_path):
    errors, _df, report = _run([_estrangeiro()], tmp_path, login_choice="EMAIL")

    assert errors == {} and report["general_errors"] == ""
    assert report["required_blank"]["CPF"] == 0


def test_com_login_por_cpf_o_cpf_segue_obrigatorio_mesmo_com_passaporte(tmp_path):
    errors, _df, _report = _run([_estrangeiro()], tmp_path, login_choice="CPF")

    assert errors[0] == "Campo obrigatório em branco: CPF"


def test_sem_cpf_e_sem_passaporte_a_linha_e_invalida_mesmo_no_login_por_email(tmp_path):
    rows = [_estrangeiro(), _row(2, CPF="")]
    rows[1]["NÚMERO PASSAPORTE"] = ""

    errors, _df, report = _run(rows, tmp_path, login_choice="EMAIL")

    assert errors == {1: "Campo obrigatório em branco: CPF"}
    assert report["required_blank"]["CPF"] == 1  # só a que não tem nem CPF nem passaporte


def test_ficha_sem_coluna_cpf_mas_com_passaporte_nao_acusa_coluna_ausente(tmp_path):
    row = _estrangeiro()
    del row["CPF"]

    errors, _df, report = _run([row], tmp_path, login_choice="EMAIL")

    assert errors == {} and report["general_errors"] == ""


def test_todos_estrangeiros_nao_geram_coluna_cpf_vazia(tmp_path):
    _errors, _df, report = _run(
        [_estrangeiro(), _estrangeiro(**{"NOME COMPLETO": "Outra Pessoa", "E-MAIL": "o@x.com"})],
        tmp_path,
        login_choice="EMAIL",
    )

    assert report["general_errors"] == ""


def test_passaporte_com_cpf_preenchido_nao_muda_nada(tmp_path):
    row = _row(1)
    row["NÚMERO PASSAPORTE (obrigatório para estrangeiro)"] = "AB123456"

    errors, df, _report = _run([row], tmp_path, login_choice="CPF")

    assert errors == {}
    assert list(df.columns) == MODEL_COLS  # passaporte é só fonte: não vira coluna de carga


# ------------------------------------------------------- formato do e-mail


@pytest.mark.parametrize(
    "email", ["a@x.com", "nome.sobrenome@empresa.com.br", "a+tag@x.io", "ÑANDÚ@x.com", "  a@x.com  "]
)
def test_emails_validos(email):
    assert "Email inválido" not in ValidationService.validate_row({"Email": email})


@pytest.mark.parametrize(
    "email",
    [
        "@x.com",
        "a@b.",
        "a b@x.com",
        "a@b@c.com",
        "a@x.com; b@x.com",
        "a@x.com, b@x.com",
        "a@x",
        "a@.com",
        "a@b..com",
        "semarroba",
        "a@b.c",
    ],
)
def test_emails_invalidos(email):
    assert "Email inválido" in ValidationService.validate_row({"Email": email})


# ------------------------------------------------------ vários arquivos


def test_arquivo_sem_uma_coluna_nao_vira_texto_nan_nem_escapa_da_validacao(tmp_path):
    a = _write([_row(1, CARGO="Analista")], tmp_path / "a.xlsx")
    sem_telefone = {k: v for k, v in _row(2).items() if k != "TELEFONE"}
    b = _write([sem_telefone], tmp_path / "b.xlsx")

    errors, df = ProcessingService.process_records_from_files([a, b])

    assert df["Cargo"].tolist() == ["ANALISTA", ""]  # antes a 2ª linha saía "NAN"
    assert not df.astype(str).apply(lambda s: s.str.upper().eq("NAN")).any().any()
    assert errors == {1: "Campo obrigatório em branco: Telefone"}  # a ficha sem a coluna não passa batida


# --------------------------------------------- planilha sem dados: por quê


def _wb(sheets, path):
    wb = openpyxl.Workbook()
    for i, (name, rows) in enumerate(sheets.items()):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = name
        for r in rows:
            ws.append(r)
    wb.save(path)
    return str(path)


HEAD = list(_row(1))


def _problem(sheets, tmp_path):
    path = _wb(sheets, tmp_path / "in.xlsx")
    errors, df = ProcessingService.process_records_from_files([path])
    assert df.empty
    return errors[path]


def test_planilha_totalmente_vazia_diz_que_esta_vazia(tmp_path):
    assert _problem({"S": []}, tmp_path) == "A planilha está vazia."


def test_so_o_cabecalho_diz_que_nao_ha_linhas_de_dados(tmp_path):
    assert _problem({"S": [HEAD]}, tmp_path) == "A planilha não tem linhas de dados abaixo do cabeçalho."


def test_nenhuma_coluna_reconhecida_lista_as_colunas_lidas(tmp_path):
    msg = _problem({"S": [["A", "B", "C"], [1, 2, 3]]}, tmp_path)

    assert msg.startswith("Nenhuma coluna da ficha foi reconhecida (colunas lidas: A, B, C)")
    assert "1ª linha" in msg


def test_titulo_acima_do_cabecalho_e_explicado_pelo_cabecalho_na_1a_linha(tmp_path):
    msg = _problem({"S": [["FICHA DE CADASTRO - SETEMBRO"], [], HEAD, list(_row(1).values())]}, tmp_path)

    assert "O cabeçalho precisa estar na 1ª linha." in msg


def test_dados_na_2a_aba_avisam_que_so_a_1a_e_lida(tmp_path):
    msg = _problem({"Leia-me": [["Preencha a aba Dados"]], "Dados": [HEAD, list(_row(1).values())]}, tmp_path)

    assert msg.endswith("Só a 1ª aba ('Leia-me') é lida.")


# ---------------------------------------------------------- pela API


def _post(client, route, files, **form):
    data = {"files[]": files, **form}
    return client.post(route, data=data, content_type="multipart/form-data")


def _upload(rows, name="fichas.xlsx"):
    return xlsx_upload(pd.DataFrame(rows), name)


def _zip_com_estrutura_minima_e_conteudo_lixo():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", "<x/>")
        z.writestr("xl/workbook.xml", "lixo")
    buf.seek(0)
    return (buf, "quebrada.xlsx")


@pytest.mark.parametrize("route", ["/api/analysis/summary", "/api/process_cadastro"])
def test_planilha_sem_dados_devolve_400_com_motivo_claro(client, route):
    resp = _post(client, route, _upload([{"A": 1, "B": 2}], "fichas.xlsx"), login_choice="CPF", fluxo="SELF")

    assert resp.status_code == 400
    body = resp.get_json()
    assert "fichas.xlsx: Nenhuma coluna da ficha foi reconhecida" in body["error"]
    # O motivo também vem por arquivo (a tela o explica sem reler o texto): validar e gerar dão a mesma resposta.
    assert body["errors"]["fichas.xlsx"].startswith("Nenhuma coluna da ficha foi reconhecida")


@pytest.mark.parametrize("route", ["/api/analysis/summary", "/api/process_cadastro"])
def test_erro_de_leitura_nao_vaza_o_caminho_do_servidor(client, tmp_path, route):
    resp = _post(client, route, _zip_com_estrutura_minima_e_conteudo_lixo(), login_choice="CPF", fluxo="SELF")

    corpo = resp.get_data(as_text=True)
    assert resp.status_code == 400
    assert "quebrada.xlsx" in corpo
    assert str(tmp_path).replace("\\", "/") not in corpo.replace("\\\\", "/").replace("\\", "/")
    assert "uploads" not in corpo


def test_arquivo_quebrado_junto_de_um_bom_aparece_no_erro_geral_do_relatorio(client):
    files = [_upload([_row(1)], "boa.xlsx"), _zip_com_estrutura_minima_e_conteudo_lixo()]

    resp = _post(client, "/api/analysis/summary", files, login_choice="CPF", fluxo="SELF")

    report = resp.get_json()["report"]
    assert resp.status_code == 200
    assert report["total_rows"] == 1
    assert report["general_errors"].startswith("quebrada.xlsx: ")  # o arquivo não some em silêncio


@pytest.mark.parametrize("route", ["/api/analysis/summary", "/api/process_cadastro"])
@pytest.mark.parametrize(
    ("form", "mensagem"),
    [({"login_choice": "XYZ"}, "Tipo de login inválido"), ({"fluxo": "FOO"}, "Fluxo inválido")],
)
def test_parametros_desconhecidos_sao_recusados(client, route, form, mensagem):
    resp = _post(client, route, _upload([_row(1)]), **form)

    assert resp.status_code == 400
    assert mensagem in resp.get_json()["error"]


@pytest.mark.parametrize("valor", ["cpf", " CPF ", ""])
def test_login_choice_ignora_caixa_e_espacos_e_vazio_vale_o_padrao(client, valor):
    resp = _post(client, "/api/process_cadastro", _upload([_row(1)]), login_choice=valor, fluxo="self")

    assert resp.status_code == 200
    out = pd.read_excel(io.BytesIO(resp.data), dtype=str).fillna("")
    assert out["Login"].tolist()[0] != ""  # antes, "cpf" gerava o arquivo com Login vazio


def test_parametros_ausentes_usam_cpf_e_self(client):
    resp = client.post(
        "/api/process_cadastro", data={"files[]": _upload([_row(1)])}, content_type="multipart/form-data"
    )

    assert resp.status_code == 200


# ------------------------------------------------------------ leitor


def test_leitor_cai_no_motor_padrao_quando_o_calamine_falha(tmp_path, monkeypatch):
    path = _write([_row(1)], tmp_path / "in.xlsx")
    original = excel_reader._read_first_sheet
    engines = []

    def fake(path, engine):
        engines.append(engine)
        if engine == "calamine":
            raise RuntimeError("calamine quebrou")
        return original(path, engine)

    monkeypatch.setattr(excel_reader, "_read_first_sheet", fake)

    df, name, count = excel_reader.read_first_sheet_as_text(path)

    assert engines == ["calamine", None]
    assert len(df) == 1 and count == 1 and name == "Sheet1"


def test_leitor_entrega_as_mesmas_celulas_nos_dois_motores(tmp_path):
    rows = [{"a": "001", "b": 12345, "c": 1.5, "d": "  x  ", "e": None, "f": "José"}]
    path = _write(rows, tmp_path / "in.xlsx")

    via_calamine, _, _ = excel_reader._read_first_sheet(path, "calamine")
    via_openpyxl, _, _ = excel_reader._read_first_sheet(path, None)

    assert via_calamine.fillna("").equals(via_openpyxl.fillna(""))


# ------------------------------------- comportamentos do pipeline vetorizado


def test_duas_colunas_para_o_mesmo_campo_vale_a_ultima(tmp_path):
    row = _row(1, **{"E-MAIL": "primeiro@x.com"})
    row["EMAIL"] = "ultimo@x.com"  # 2ª coluna que também vira Email

    _errors, df, _report = _run([row], tmp_path)

    assert df["Email"].tolist() == ["ULTIMO@X.COM"]


def test_linha_que_repete_o_cabecalho_e_descartada_e_a_numeracao_segue_o_excel(tmp_path):
    rows = [_row(1), {k: k for k in _row(1)}, _row(3, TELEFONE="")]  # a 2ª linha de dados repete o cabeçalho

    errors, df, report = _run(rows, tmp_path)

    assert len(df) == 2
    assert list(report["line_errors"]) == ["Linha 4"]  # a linha 3 do Excel era o cabeçalho repetido


def test_drop_header_like_rows_usa_o_limite_de_40_por_cento_das_colunas():
    df = pd.DataFrame(
        {"A": ["a", "x", "a"], "B": ["b", "b", "q"], "C": ["c", "y", "z"], "D": ["d", "y", "z"], "E": ["e", "y", "z"]}
    )

    kept = ProcessingService._drop_header_like_rows(df)

    # linha 0: 5/5 iguais ao cabeçalho; linha 1: 1/5 (20%); linha 2: 1/5 -> só a 0 sai
    assert kept.index.tolist() == [1, 2]


def test_map_unique_calcula_cada_valor_uma_vez_e_aceita_valor_nao_hasheavel():
    from backend.services.processing_service import _map_unique

    calls = []

    def fn(v):
        calls.append(v)
        return str(v).upper()

    assert _map_unique(["a", "b", "a", "a", ["x"], ["x"]], fn) == ["A", "B", "A", "A", "['X']", "['X']"]
    assert calls == ["a", "b", ["x"], ["x"]]  # "a" só uma vez; a lista não é hasheável e é calculada direto


def test_varios_arquivos_com_colunas_em_ordem_diferente_juntam_as_linhas_na_ordem_dos_arquivos(tmp_path):
    a = _write([_row(1)], tmp_path / "a.xlsx")
    ordem_invertida = {k: v for k, v in reversed(list(_row(2).items()))}
    b = _write([ordem_invertida], tmp_path / "b.xlsx")

    errors, df = ProcessingService.process_records_from_files([a, b])

    assert errors == {}
    assert df["NomeCompleto"].tolist() == ["PESSOA NUMERO1", "PESSOA NUMERO2"]


def test_rotulos_de_linha_so_existem_para_linhas_com_erro(tmp_path):
    _errors, df, _report = _run([_row(1), _row(2, TELEFONE=""), _row(3)], tmp_path)

    # df.attrs é copiado a cada operação do pandas: um rótulo por linha chegou a custar 1/3 do tempo de exportar
    assert df.attrs["row_labels"] == {1: "Linha 3"}


def test_volume_moderado_processa_em_tempo_razoavel(tmp_path):
    import time

    rows = [_row(i) | {"NOME COMPLETO": f"Pessoa Numero {i}", "E-MAIL": f"p{i}@x.com"} for i in range(1, 3001)]
    path = _write(rows, tmp_path / "grande.xlsx")

    start = time.time()
    errors, df = ProcessingService.process_records_from_files([path])

    assert len(df) == 3000 and errors == {}
    assert time.time() - start < 15  # folga enorme (leva ~0,4 s): só pega regressão para laço linha a linha lento
