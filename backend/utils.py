# backend/utils.py
import re
import unicodedata

def normalize_text(s):
    if s is None:
        return ""
    return str(s).strip()

def upper_no_accents(s):
    s = normalize_text(s).upper()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")
    return s

def limpar_cpf_raw(cpf):
    """retorna apenas digitos do CPF"""
    if cpf is None:
        return ""
    s = str(cpf)
    return re.sub(r"\D", "", s)

def format_cpf_for_output(cpf_digits):
    """formata com traço antes dos 2 ultimos digitos: 12345678901 -> 123456789-01"""
    if not cpf_digits:
        return ""
    s = re.sub(r"\D", "", str(cpf_digits))
    if len(s) == 11:
        return f"{s[:-2]}-{s[-2:]}"
    return s


def separar_nome_sobrenome(nome_completo: str) -> tuple[str, str]:
    """Separa um nome completo em (nome, sobrenome) respeitando regras de
    nomes e sobrenomes compostos, partículas e sufixos geracionais.

    Exemplos:
    >>> separar_nome_sobrenome("João Pedro da Silva")
    ('João Pedro', 'da Silva')

    >>> separar_nome_sobrenome("Maria Eduarda dos Santos Neto")
    ('Maria Eduarda', 'dos Santos Neto')

    >>> separar_nome_sobrenome("Luiz Inácio Lula da Silva")
    ('Luiz Inácio Lula', 'da Silva')

    >>> separar_nome_sobrenome("Ana Clara Alves Neto")
    ('Ana Clara', 'Alves Neto')

    >>> separar_nome_sobrenome("José Silva")
    ('José', 'Silva')

    >>> separar_nome_sobrenome("Carlos Alberto de Oliveira Filho")
    ('Carlos Alberto', 'de Oliveira Filho')
    """
    if not nome_completo:
        return "", ""

    partes = [p for p in str(nome_completo).strip().split() if p]
    n = len(partes)
    if n == 1:
        return partes[0], ""

    particulas = {
        "da", "de", "do", "das", "dos", "di", "della", "del", "dela",
        "van", "von",
    }

    sufixos_geracionais = {
        "neto",
        "netto",
        "filho",
        "filha",
        "junior",
        "júnior",
        "jr",
        "sobrinho",
    }

    def is_sufixo(palavra: str) -> bool:
        return palavra.casefold() in sufixos_geracionais

    def is_particula(palavra: str) -> bool:
        return palavra.casefold() in particulas

    sobrenome: list[str] = []

    # Casos simples por quantidade de partes, antes de lidar com partículas/sufixos
    if n == 2:
        # Pedro Silva -> Nome: Pedro | Sobrenome: Silva
        return partes[0], partes[1]

    if n == 3 and not is_sufixo(partes[-1]):
        # Andre Gomes Lima -> Nome: Andre | Sobrenome: Gomes Lima
        # Giovanna Almeida Lima -> Nome: Giovanna | Sobrenome: Almeida Lima
        return partes[0], " ".join(partes[1:])

    # A partir daqui, tratamos casos gerais (inclui 3 com sufixo, 4+ etc.)
    # Regra base: sempre começar do fim
    sobrenome.insert(0, partes[-1])
    i = n - 2

    # Se o último termo for sufixo geracional (NETO, FILHO, etc.) e ainda só
    # houver uma palavra no sobrenome, garantir ao menos 2 termos.
    if is_sufixo(sobrenome[-1]) and i >= 0 and len(sobrenome) < 2:
        sobrenome.insert(0, partes[i])
        i -= 1

    # Para nomes com 4+ partes (incluindo os com sufixo), usar DOIS últimos
    # termos como base do sobrenome se ainda houver apenas 1 termo até aqui.
    if n >= 4 and i >= 0 and len(sobrenome) == 1:
        sobrenome.insert(0, partes[i])
        i -= 1

    # Puxar partículas imediatamente anteriores para o sobrenome
    while i >= 0 and is_particula(partes[i]):
        sobrenome.insert(0, partes[i])
        i -= 1

    nome = " ".join(partes[: i + 1]).strip()
    sobrenome_str = " ".join(sobrenome).strip()

    if not nome and partes:
        nome = partes[0]
        sobrenome_str = " ".join(partes[1:]).strip()

    return nome, sobrenome_str
