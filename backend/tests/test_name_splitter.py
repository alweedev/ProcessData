"""Separação do nome completo em Nome e Sobrenome (nomes fictícios, só a estrutura importa)."""

import json

import pytest

from backend.shared.name_splitter import (
    ACCEPTED_WEIGHT,
    CONFIDENT,
    DOUBTFUL,
    EDITED_WEIGHT,
    MAX_FIELD_LEN,
    NameVocabulary,
    get_vocabulary,
    split_full_name,
    suggest_abbreviation,
)


@pytest.fixture
def vocab():
    return NameVocabulary(None)  # só a base embutida, sem arquivo


@pytest.mark.parametrize(
    ("entrada", "nome", "sobrenome"),
    [
        # partícula é do sobrenome que vem depois
        ("Maria Clara da Silva Santos", "MARIA CLARA", "DA SILVA SANTOS"),
        ("Ana Paula de Souza", "ANA PAULA", "DE SOUZA"),
        ("Pedro Henrique van der Berg", "PEDRO HENRIQUE", "VAN DER BERG"),
        ("Luiz Fernando de Souza e Silva", "LUIZ FERNANDO", "DE SOUZA E SILVA"),
        ("Fulano de Tal", "FULANO", "DE TAL"),
        # sufixo gruda no sobrenome anterior
        ("João Carlos dos Santos Filho", "JOAO CARLOS", "DOS SANTOS FILHO"),
        ("João Silva Neto", "JOAO", "SILVA NETO"),
        # nome composto, com e sem partícula
        ("Carlos Eduardo Lima", "CARLOS EDUARDO", "LIMA"),
        ("José Maria Pereira", "JOSE MARIA", "PEREIRA"),
        ("Maria das Graças Xavier", "MARIA DAS GRACAS", "XAVIER"),
        ("Maria de Fátima Lima", "MARIA DE FATIMA", "LIMA"),
        ("Carlos-Eduardo Lima", "CARLOS-EDUARDO", "LIMA"),
        # "Gabriel" é nome próprio e também sobrenome: aqui é nome (mesmo padrão de "X Gabriel Almeida Y")
        ("Rodrigo Gabriel Almeida Herrera", "RODRIGO GABRIEL", "ALMEIDA HERRERA"),
        ("Juan Pablo Perez Gonzalez", "JUAN PABLO", "PEREZ GONZALEZ"),
        ("Ana Silva", "ANA", "SILVA"),
    ],
)
def test_separa_nome_e_sobrenome_como_no_documento(vocab, entrada, nome, sobrenome):
    split = split_full_name(entrada, vocab)

    assert (split.nome, split.sobrenome) == (nome, sobrenome)
    assert split.confidence == CONFIDENT


def test_acento_apostrofo_e_espacos_sao_normalizados(vocab):
    split = split_full_name("  Jose   D'Avila  Nandú ", vocab)

    assert (split.nome, split.sobrenome) == ("JOSE", "DAVILA NANDU")
    assert split.confidence == DOUBTFUL  # duas palavras fora do vocabulário no meio e no fim: a divisão é ambígua


@pytest.mark.parametrize(
    ("entrada", "motivo"),
    [
        ("Madonna", "só uma palavra"),
        ("Pessoa 123", "contém número"),
        ("Ana Paula", "sem sobrenome reconhecido"),
        ("Carlos Junior", "sufixo, mas não há sobrenome antes"),
        ("Maria Clara Neto", "sufixo, mas não há sobrenome antes"),
        ("Beltrano Zeferino Quintela Bragança", "ambígua"),  # nenhuma palavra conhecida
        ("Ana Maria Clara Beatriz Helena Souza Lima", "muito longo"),
    ],
)
def test_casos_duvidosos_vao_para_conferencia_com_o_motivo(vocab, entrada, motivo):
    split = split_full_name(entrada, vocab)

    assert split.confidence == DOUBTFUL
    assert any(motivo in r for r in split.reasons), split.reasons


def test_nome_vazio_nao_quebra(vocab):
    assert split_full_name("", vocab).nome == ""
    assert split_full_name(None, vocab).sobrenome == ""
    assert split_full_name("???", vocab).confidence == DOUBTFUL


def test_particula_no_inicio_e_duvidosa(vocab):
    assert split_full_name("De Souza João", vocab).confidence == DOUBTFUL


def test_o_resultado_e_deterministico(vocab):
    a = split_full_name("Maria Clara da Silva Santos", vocab)
    b = split_full_name("MARIA CLARA DA SILVA SANTOS", vocab)

    assert a == b


# --------------------------------------------------------------- vocabulário aprendido


def test_confirmacao_ensina_o_vocabulario_e_a_duvida_some(tmp_path):
    vocab = NameVocabulary(str(tmp_path / "v.json"))
    nome = "Beltrano Zeferino Quintela Bragança"
    assert split_full_name(nome, vocab).confidence == DOUBTFUL

    for _ in range(2):  # o usuário CORRIGE à mão a mesma divisão duas vezes
        vocab.learn("BELTRANO ZEFERINO", "QUINTELA BRAGANCA", weight=EDITED_WEIGHT)

    depois = split_full_name(nome, vocab)
    assert (depois.nome, depois.sobrenome) == ("BELTRANO ZEFERINO", "QUINTELA BRAGANCA")
    assert depois.confidence == CONFIDENT


def test_aceitar_a_sugestao_ensina_menos_que_corrigir(tmp_path):
    aceito = NameVocabulary(str(tmp_path / "a.json"))
    corrigido = NameVocabulary(str(tmp_path / "c.json"))

    aceito.learn("BELTRANO", "QUINTELA", weight=ACCEPTED_WEIGHT)
    corrigido.learn("BELTRANO", "QUINTELA", weight=EDITED_WEIGHT)

    assert aceito.given_count("BELTRANO") == 1 and corrigido.given_count("BELTRANO") == 3


def test_um_aceitar_todas_apressado_nao_vira_certeza(tmp_path):
    # aceitar a mesma sugestão duas vezes ainda deixa a divisão em dúvida (só a correção manual pesa mais)
    vocab = NameVocabulary(str(tmp_path / "v.json"))
    for _ in range(2):
        vocab.learn("BELTRANO ZEFERINO", "QUINTELA BRAGANCA", weight=ACCEPTED_WEIGHT)

    assert split_full_name("Beltrano Zeferino Quintela Bragança", vocab).confidence == DOUBTFUL


def test_so_palavras_soltas_sao_gravadas_nunca_o_nome_completo(tmp_path):
    path = tmp_path / "v.json"
    vocab = NameVocabulary(str(path))

    vocab.learn("BELTRANO ZEFERINO", "DA QUINTELA 123 X BRAGANCA")

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved == {"given": {"BELTRANO": 1, "ZEFERINO": 1}, "surname": {"QUINTELA": 1, "BRAGANCA": 1}}
    # partícula, número e inicial solta não entram; nada de "BELTRANO ZEFERINO" junto


def test_vocabulario_aprendido_sobrevive_a_reiniciar(tmp_path):
    path = str(tmp_path / "v.json")
    NameVocabulary(path).learn("BELTRANO", "QUINTELA")

    novo = NameVocabulary(path)

    assert novo.given_count("BELTRANO") == 1
    assert novo.surname_count("QUINTELA") == 1


def test_arquivo_de_vocabulario_corrompido_nao_derruba_o_cadastro(tmp_path):
    path = tmp_path / "v.json"
    path.write_text("{isso não é json", encoding="utf-8")

    vocab = NameVocabulary(str(path))

    assert split_full_name("Maria Clara da Silva Santos", vocab).nome == "MARIA CLARA"  # segue na base embutida


def test_falha_ao_gravar_nao_levanta_erro(tmp_path):
    bloqueio = tmp_path / "arquivo"
    bloqueio.write_text("x")
    vocab = NameVocabulary(str(bloqueio / "sub" / "v.json"))  # a pasta "arquivo" é um arquivo: não dá para criar

    assert vocab.learn("BELTRANO", "QUINTELA") is False
    assert vocab.given_count("BELTRANO") == 1  # ainda vale nesta sessão


def test_uma_confirmacao_nao_desmonta_o_que_a_base_embutida_sabe(tmp_path):
    vocab = NameVocabulary(str(tmp_path / "v.json"))
    vocab.learn("SILVA", "MARIA")  # engano: inverteu nome e sobrenome uma vez

    assert split_full_name("Maria Clara da Silva Santos", vocab).nome == "MARIA CLARA"


def test_get_vocabulary_usa_o_arquivo_configurado(tmp_path):
    from backend.core.config import settings

    a = get_vocabulary()

    assert a.path == settings.NAME_VOCAB_FILE
    assert get_vocabulary() is a  # um por caminho


# --------------------------------------------------------------- tamanho


def test_abreviacao_sugerida_cabe_em_20_e_preserva_o_primeiro_nome_e_o_ultimo_sobrenome():
    nome, sobrenome = suggest_abbreviation("MARIA CLARA BEATRIZ HELENA", "FERNANDES DE ALBUQUERQUE CAVALCANTI")

    assert (nome, sobrenome) == ("MARIA CLARA B H", "F DE A CAVALCANTI")
    assert len(nome) <= MAX_FIELD_LEN and len(sobrenome) <= MAX_FIELD_LEN


def test_quem_ja_cabe_nao_e_abreviado():
    assert suggest_abbreviation("ANA PAULA", "DE SOUZA") == ("ANA PAULA", "DE SOUZA")


def test_abreviacao_mantem_particula_e_sufixo():
    _, sobrenome = suggest_abbreviation("ANA", "DOS SANTOS FILHO DE ALBUQUERQUE MARANHAO")

    assert sobrenome == "DOS S FILHO DE A MARANHAO"
    assert len(sobrenome) > MAX_FIELD_LEN  # não cabe nem abreviando: quem confere edita à mão


def test_nome_que_nao_cabe_nem_abreviando_continua_acima_do_limite():
    nome, _ = suggest_abbreviation("PNEUMOULTRAMICROSCOPICOSSILICOVULCANO", "SILVA")

    assert len(nome) > MAX_FIELD_LEN
