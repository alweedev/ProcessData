from backend.shared.cpf_mask import mascarar_cpf


def test_mostra_so_os_seis_digitos_do_meio():
    assert mascarar_cpf("12345678909") == "***.456.789-**"


def test_aceita_pontuacao():
    assert mascarar_cpf("123.456.789-09") == "***.456.789-**"


def test_restaura_zero_a_esquerda_perdido_pelo_excel():
    assert mascarar_cpf("1234567890") == "***.345.678-**"


def test_vazio_ou_invalido_devolve_vazio():
    assert mascarar_cpf("") == ""
    assert mascarar_cpf(None) == ""
    assert mascarar_cpf("abc") == ""
    assert mascarar_cpf("123456789012") == ""
