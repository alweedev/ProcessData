"""Separa o nome completo em Nome (todos os nomes próprios) e Sobrenome (todos os sobrenomes).

As companhias aéreas exigem o nome como no documento: "Maria Clara da Silva Santos" precisa virar
Nome=MARIA CLARA e Sobrenome=DA SILVA SANTOS, não "primeira e última palavra". Isso combina três coisas:

1. **Regras.** Partícula (DA, DE, DOS, VAN DER...) faz parte do sobrenome que vem depois; sufixo (FILHO, JUNIOR,
   NETO...) gruda no sobrenome anterior; "Maria das Graças" e similares são nome composto.
2. **Vocabulário.** Uma lista de nomes próprios e outra de sobrenomes, com frequência, diz de que lado da divisão
   cada palavra costuma ficar. Começa embutida e cresce com as conferências do usuário (`NameVocabulary.learn`).
3. **Confiança.** Cada divisão sai com "alta" ou "baixa". Só as de baixa vão para a tela de conferir.

Tudo local e determinístico: nenhum nome de passageiro sai da máquina, e o vocabulário aprendido guarda só
palavras soltas com contagem (nunca o nome completo nem o CPF).
"""

from __future__ import annotations

import json
import logging
import math
import os
import tempfile
import threading
from collections import Counter
from dataclasses import dataclass

from backend.core.config import settings
from backend.shared.text_utils import sanitize_output_text

logger = logging.getLogger(__name__)

MAX_FIELD_LEN = 20  # limite de Nome e de Sobrenome na plataforma
CONFIDENT = "alta"
DOUBTFUL = "baixa"

# Partículas: fazem parte do sobrenome que vem depois delas ("DA SILVA", "VAN DER BERG", "DE LA CRUZ").
PARTICLES = frozenset(
    "DA DE DO DAS DOS DI DU DEL DELLA DELLE DELLI LA LE LOS LAS EL AL VAN VON DER DEN TER TEN Y E BEN BIN BINT".split()
)
# Sufixos: identificam a geração e ficam no sobrenome ("SANTOS FILHO", "SILVA NETO", "LIMA JUNIOR").
SUFFIXES = frozenset(
    "FILHO FILHA JUNIOR JR NETO NETA NETTO SOBRINHO SOBRINHA BISNETO BISNETA SEGUNDO TERCEIRO II III IV".split()
)
# Nomes compostos com partícula, muito comuns no Brasil: a partícula faz parte do NOME, não do sobrenome.
COMPOUND_GIVEN = frozenset(
    {
        ("MARIA", "DA", "GRACA"),
        ("MARIA", "DAS", "GRACAS"),
        ("MARIA", "DA", "GLORIA"),
        ("MARIA", "DA", "PENHA"),
        ("MARIA", "DA", "CONCEICAO"),
        ("MARIA", "DA", "LUZ"),
        ("MARIA", "DA", "PAZ"),
        ("MARIA", "DE", "FATIMA"),
        ("MARIA", "DE", "LOURDES"),
        ("MARIA", "DO", "CARMO"),
        ("MARIA", "DO", "SOCORRO"),
        ("MARIA", "DO", "ROSARIO"),
        ("MARIA", "DAS", "DORES"),
        ("MARIA", "DOS", "ANJOS"),
        ("MARIA", "DOS", "SANTOS"),
        ("ANA", "DA", "LUZ"),
        ("JOSE", "DE", "JESUS"),
    }
)

# --- Vocabulário embutido -------------------------------------------------------------------------------------
# Sem acento e em maiúsculas (é como o texto chega ao separador). "PALAVRA:peso" reduz o peso de quem também
# aparece do outro lado (GABRIEL é nome próprio; como sobrenome é raro). É só o ponto de partida: as conferências
# do usuário ensinam o resto. Nomes fora daqui não são erro: o separador usa a posição e marca "baixa" se duvidar.
_SEED_WEIGHT = 20

_SEED_GIVEN = """
MARIA ANA FRANCISCA ANTONIA ADRIANA JULIANA MARCIA FERNANDA PATRICIA ALINE SANDRA CAMILA AMANDA BRUNA JESSICA
LETICIA JULIA LUCIANA VANESSA MARIANA GABRIELA VERA VITORIA LARISSA CLAUDIA CRISTINA SIMONE ROSANGELA DANIELA
RAQUEL RENATA CAROLINA BEATRIZ ISABELA ISABEL LAURA LUCIA LUIZA LUISA HELENA SOFIA SOPHIA ALICE VALENTINA MANUELA
LIVIA CLARA PAULA PAOLA THAIS TATIANA TEREZA TERESA REGINA ROBERTA ROSA SILVIA SUELI SONIA SHIRLEY VIVIAN ELAINE
ELIANE ELISABETE ELIZABETH EDNA ERIKA EVELYN FABIANA FLAVIA GISELE GLAUCIA GRACA IRENE IVONE JAQUELINE JOANA
JOSEFA JOSEFINA KARINA KATIA LILIAN LIDIA LEILA MARLENE MICHELE MONICA NATALIA NATHALIA NEUSA NICOLE NILZA NOEMI
OLIVIA PRISCILA RAFAELA ROSANE ROSELI SABRINA SAMARA SARA SARAH SELMA TANIA VALERIA VIVIANE YASMIN YARA ZILDA
ALESSANDRA ANGELA ANGELICA ANITA APARECIDA BARBARA BIANCA BRENDA CAROLINE CECILIA CELIA CIBELE CLARICE DEBORA
DENISE DIANA EDUARDA ELISA EMANUELLE EMILY ESTER EVA FATIMA FELIPA GEOVANA GIOVANA HELOISA INES IRIS IRACEMA
JANAINA JOELMA KELLY LAIS LARA LORENA LUANA LUCIANE MAIARA MARCELA MARGARIDA MARILENE MARINA MARISA MAYARA MELISSA
MIRIAM NAIARA NAYARA NADIA NATASHA PAMELA PIETRA ROSILENE SAMANTA STEFANY STEPHANIE SUZANA TAINA TAMIRES THALITA
VANDA VILMA WANDA CARMEN CARMELITA CONCEICAO DOLORES DORA ELENA EMILIA ESPERANCA FRANCINE GABRIELLE GLORIA GUADALUPE
JOSE JOAO ANTONIO FRANCISCO CARLOS PAULO PEDRO LUCAS LUIZ LUIS MARCOS GABRIEL RAFAEL DANIEL MARCELO BRUNO EDUARDO
FELIPE RODRIGO MANOEL MANUEL GUSTAVO ANDRE FERNANDO FABIO LEONARDO GILBERTO GUILHERME HENRIQUE JORGE JULIO LEANDRO
MATEUS MATHEUS MAURICIO MIGUEL NELSON RICARDO ROBERTO RONALDO SERGIO THIAGO TIAGO VINICIUS VICTOR VITOR WILLIAM
WELLINGTON ADRIANO ALAN ALBERTO ALEX ALEXANDRE ALEJANDRO ALFREDO ALVARO ANDERSON ANGELO ANTONY ARTHUR ARTUR AUGUSTO
BENEDITO BENJAMIN CAIO CAIQUE CASSIO CESAR CLAUDIO CLEBER CRISTIANO DAVI DAVID DIEGO DIOGO DOUGLAS EDSON EDGAR ELIAS
EMERSON ENZO ERICK ERIC EVERTON FABRICIO FAUSTO FLAVIO GERALDO GERSON GIOVANNI GIOVANI GREGORIO HEITOR HUGO IGOR
ISAAC ITALO IVAN JAIME JAIR JEAN JEFERSON JEFFERSON JESUS JOAQUIM JONAS JONATHAN JOSUE JUAN KAUA KEVIN LAURO LEO
LEONEL LINCOLN LORENZO LUAN LUCIANO MARIO MARLON MARTIN MAX MICAEL MOISES MURILO NATAN NICOLAS NOEL OSCAR OSVALDO
OTAVIO PABLO PATRICK RAUL RENAN RENATO REINALDO RENE ROBSON RODOLFO ROGERIO ROMULO RUAN RUBENS RUI SAMUEL SANTIAGO
SAULO SEBASTIAO SIDNEY SILVIO STEFAN TALES TARCISIO TELMO THEO TOMAS VALDIR VALTER VANDERLEI VICENTE WAGNER WALTER
WALLACE WESLEY YURI ADEMIR ALEXSANDRO AMARO ANDRES ARMANDO BERNARDO CAMILO CRISTIAN CRISTOVAO DANILO DARIO DENIS
DIONISIO EDMUNDO ELTON EMANUEL ENRIQUE ERNESTO ESTEVAO EUGENIO EZEQUIEL FELIX FILIPE FREDERICO GENIVAL HAROLDO HELIO
HERALDO HERMES HORACIO IRINEU ISMAEL JACKSON JOEL JORDAO JULIANO KLEBER LAERTE LEVI LUCIO MARCIO MARCELINO MATIAS
MAURO NILSON NILTON ORLANDO OSMAR PEDRINHO RAIMUNDO RAFAELA RAMON ROBERT ROMEU RONALD SALVADOR SAMIR SIMAO TEODORO
THOMAS TIMOTEO ULISSES VALMIR VLADIMIR WALDEMAR WASHINGTON WELLINGTON WILSON YAGO ZACARIAS
"""

_SEED_SURNAMES = """
SILVA SANTOS OLIVEIRA SOUZA SOUSA RODRIGUES FERREIRA ALVES PEREIRA LIMA GOMES COSTA RIBEIRO MARTINS CARVALHO ALMEIDA
LOPES SOARES FERNANDES VIEIRA BARBOSA ROCHA DIAS NASCIMENTO ANDRADE MOREIRA NUNES MARQUES MACHADO MENDES FREITAS
CARDOSO RAMOS GONCALVES SANTANA TEIXEIRA CAVALCANTI MONTEIRO ARAUJO CORREIA MOURA PINTO MIRANDA BARROS CAMPOS MELO
MELLO CASTRO BORGES AZEVEDO CUNHA CAMARGO TAVARES PIRES FARIAS CAETANO REIS GUIMARAES LEITE MEDEIROS BATISTA BRAGA
DUARTE FONSECA PAIVA ALBUQUERQUE BEZERRA MAGALHAES XAVIER COELHO SIQUEIRA VASCONCELOS BASTOS BRITO CARNEIRO
ABREU AGUIAR AMARAL ANDRE:5 ASSIS ASSUNCAO BARRETO BENTO BERNARDES BUENO CABRAL CALDAS CARVALHAL CHAVES CORDEIRO
COUTINHO CRUZ DAMASCENO DANTAS DOMINGUES ESTEVES EVANGELISTA FACHINI FALCAO FERRAZ FIGUEIREDO FRANCA GALVAO GAMA
GASPAR GUEDES HENRIQUES JARDIM LACERDA LAGE LAZARO LEAL LEMES LOBO LOURENCO LUZ MACEDO MAIA MALTA MARINHO MATOS
MATTOS MATIAS MOTA NEVES NOBREGA OLIVEIRA PACHECO PADILHA PAULINO PENA PERES PESSOA PIMENTA PIMENTEL PINHEIRO
PORTO PRADO QUEIROZ QUINTANILHA RABELO REZENDE RESENDE ROSA:6 SALGADO SAMPAIO SANTIAGO:6 SEVERO SILVEIRA SIMOES
SOBRAL TAVEIRA TOLEDO TRINDADE VALENTE VALADARES VARGAS VELOSO VERAS VIANA VIDAL VILELA VILLAS
GARCIA MARTINEZ GONZALEZ HERNANDEZ LOPEZ SANCHEZ RAMIREZ TORRES FLORES DIAZ ROMERO MORALES ORTIZ CASTILLO JIMENEZ RUIZ
ALVAREZ MORENO MUNOZ GUTIERREZ ROJAS DOMINGUEZ HERRERA MEDINA AGUILAR CASTANO RIOS SALAZAR CARRILLO PEREZ
SMITH JOHNSON BROWN MILLER MULLER SCHMIDT SCHNEIDER FISCHER WEBER MEYER ROSSI RUSSO FERRARI ESPOSITO BIANCHI ROMANO
COLOMBO RICCI MARINO GRECO BRUNO:3 DEBRITO KIM LEE PARK CHEN WANG WILLIAMS JONES DAVIS ANDERSON:4 TAYLOR THOMAS:3
MOORE JACKSON:4 MARTIN:3 WHITE HARRIS CLARK LEWIS ROBINSON WALKER YOUNG ALLEN KING WRIGHT SCOTT GREEN BAKER
NETO:10 JUNIOR:10 FILHO:10 SOBRINHO:10
GABRIEL:2 PEDRO:2 DANIEL:2 LUCAS:2 ANTONIO:2 RAFAEL:2 SAMUEL:2 MATEUS:2 MARCOS:2 JOAO:1 JOSE:1 PAULO:2 ELIAS:2
NICOLAU:6 DAVI:2 ISAAC:2 JORDAN:4 WILSON:4 EDUARDO:2 FELIPE:2 VICENTE:3 LEONARDO:2 RICARDO:2 ROBERTO:2 MAURICIO:2
"""


def _parse_seed(text: str) -> Counter:
    seed: Counter = Counter()
    for entry in text.split():
        word, _, weight = entry.partition(":")
        seed[word] += int(weight) if weight else _SEED_WEIGHT
    return seed


# --- Vocabulário aprendido --------------------------------------------------------------------------------------
class NameVocabulary:
    """Contagem de palavras como nome próprio e como sobrenome: base embutida + o que o usuário confirmou."""

    def __init__(self, path: str | None = None):
        self.path = path
        self._lock = threading.Lock()
        self._given = _parse_seed(_SEED_GIVEN)
        self._surname = _parse_seed(_SEED_SURNAMES)
        self._learned_given: Counter = Counter()
        self._learned_surname: Counter = Counter()
        self._load()

    def _load(self) -> None:
        if not self.path or not os.path.exists(self.path):
            return
        try:
            with open(self.path, encoding="utf-8") as fh:
                data = json.load(fh)
            self._learned_given = Counter({str(k): int(v) for k, v in data.get("given", {}).items()})
            self._learned_surname = Counter({str(k): int(v) for k, v in data.get("surname", {}).items()})
        except Exception:
            # Arquivo corrompido não pode derrubar o cadastro: segue só com a base embutida.
            logger.warning("Vocabulário de nomes ilegível (%s); usando só a base embutida.", self.path)

    def given_count(self, word: str) -> int:
        return self._given[word] + self._learned_given[word]

    def surname_count(self, word: str) -> int:
        return self._surname[word] + self._learned_surname[word]

    def p_given(self, word: str) -> float | None:
        """Probabilidade de a palavra ser nome próprio (0..1), ou None se o vocabulário não a conhece."""
        given, surname = self.given_count(word), self.surname_count(word)
        if given + surname == 0:
            return None
        return (given + 0.5) / (given + surname + 1)

    def learn(self, nome: str, sobrenome: str, weight: int = 1) -> bool:
        """Uma divisão confirmada pelo usuário (ver `learn_many`)."""
        return self.learn_many([(nome, sobrenome, weight)])

    def learn_many(self, decisions) -> bool:
        """Registra divisões confirmadas pelo usuário, ``(nome, sobrenome, peso)`` cada: cada palavra do Nome conta como
        nome próprio e cada palavra do Sobrenome, como sobrenome (partículas, iniciais e números não contam). Só palavras
        soltas são guardadas, e o arquivo é gravado uma vez só.

        O peso: aceitar a sugestão é um sinal fraco (pode ter sido um "aceitar todas" apressado, e aprender um erro o
        transformaria em certeza); editar é correção explícita e vale mais (`ACCEPTED_WEIGHT` / `EDITED_WEIGHT`).

        Melhor esforço: falha ao gravar não impede o cadastro (devolve False)."""

        def words(text: str) -> list[str]:
            found = []
            for word in sanitize_output_text(text, None).split():
                word = word.replace("-", " ").split()[0] if "-" in word else word
                if len(word) >= 2 and word not in PARTICLES and not any(ch.isdigit() for ch in word):
                    found.append(word)
            return found

        with self._lock:
            for nome, sobrenome, weight in decisions:
                for word in words(nome):
                    self._learned_given[word] += weight
                for word in words(sobrenome):
                    self._learned_surname[word] += weight
            return self._save()

    def _save(self) -> bool:
        if not self.path:
            return True
        try:
            folder = os.path.dirname(self.path)
            if folder:
                os.makedirs(folder, exist_ok=True)
            fd, tmp_path = tempfile.mkstemp(dir=folder or None, suffix=".tmp")
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                json.dump(
                    {"given": dict(self._learned_given), "surname": dict(self._learned_surname)}, fh, ensure_ascii=False
                )
            os.replace(tmp_path, self.path)
            return True
        except Exception:
            logger.warning("Não foi possível gravar o vocabulário de nomes em %s.", self.path)
            return False


_VOCABULARIES: dict[str, NameVocabulary] = {}
_VOCABULARIES_LOCK = threading.Lock()


def get_vocabulary() -> NameVocabulary:
    """O vocabulário do arquivo configurado (um por caminho, carregado uma vez por processo)."""
    path = settings.NAME_VOCAB_FILE
    with _VOCABULARIES_LOCK:
        if path not in _VOCABULARIES:
            _VOCABULARIES[path] = NameVocabulary(path)
        return _VOCABULARIES[path]


# --- Separação -------------------------------------------------------------------------------------------------
@dataclass(frozen=True)
class NameSplit:
    nome: str
    sobrenome: str
    confidence: str  # CONFIDENT ("alta": segue sozinho) ou DOUBTFUL ("baixa": vai para a conferência)
    reasons: tuple[str, ...] = ()


@dataclass(frozen=True)
class _Unit:
    text: str  # o que sai no resultado ("DA SILVA", "SANTOS FILHO")
    core: str  # a palavra que vale para o vocabulário ("SILVA", "SANTOS")
    kind: str = "word"  # word | particle | suffixed | compound_given


ACCEPTED_WEIGHT = 1  # peso de aprendizado de uma sugestão aceita como veio
EDITED_WEIGHT = 3  # ...e de uma que o usuário corrigiu à mão

_PRIOR_GIVEN_COUNT = {1: 0.60, 2: 0.32, 3: 0.07, 4: 0.01}  # quantas unidades de nome costumam vir antes do sobrenome
_CONFIDENCE_THRESHOLD = 0.80  # probabilidade mínima da melhor divisão para seguir sem conferência
_MAX_UNITS_CONFIDENT = 6  # nome com mais unidades que isso é incomum: vai para conferência


def _p_given_word(word: str, vocab: NameVocabulary) -> float | None:
    """Vocabulário para uma palavra; hifenizada ("CARLOS-EDUARDO") vale pela 1ª parte."""
    return vocab.p_given(word.split("-")[0] if "-" in word else word)


def _build_units(words: list[str], vocab: NameVocabulary, suffix_min_units: int = 2) -> list[_Unit]:
    """`suffix_min_units`: quantas unidades precisam existir antes para um sufixo grudar na última (2 = há pelo
    menos um nome antes do sobrenome; 1 serve para texto que já é só o sobrenome)."""
    units: list[_Unit] = []
    i = 0
    # "MARIA DAS GRACAS", "MARIA DE FATIMA"...: a partícula é do nome, não do sobrenome.
    if len(words) >= 3 and tuple(words[:3]) in COMPOUND_GIVEN:
        units.append(_Unit(" ".join(words[:3]), words[0], "compound_given"))
        i = 3
    while i < len(words):
        word = words[i]
        if word in PARTICLES:
            j = i
            while j < len(words) and words[j] in PARTICLES:
                j += 1
            if j < len(words):
                units.append(_Unit(" ".join(words[i : j + 1]), words[j], "particle"))
                i = j + 1
            else:  # partícula sobrando no fim: gruda no que veio antes (a conferência vai apontar)
                tail = " ".join(words[i:j])
                if units:
                    last = units[-1]
                    units[-1] = _Unit(
                        f"{last.text} {tail}", last.core, last.kind if last.kind != "word" else "suffixed"
                    )
                else:
                    units.append(_Unit(tail, words[-1], "particle"))
                i = j
        elif word in SUFFIXES and len(units) >= suffix_min_units and _suffix_attaches_to(units[-1], vocab):
            last = units[-1]
            units[-1] = _Unit(f"{last.text} {word}", last.core, "suffixed")
            i += 1
        else:
            units.append(_Unit(word, word))
            i += 1
    return units


def _suffix_attaches_to(unit: _Unit, vocab: NameVocabulary) -> bool:
    """Sufixo grudado num sobrenome ("SANTOS FILHO"), mas não num nome próprio ("MARIA CLARA" + "NETO")."""
    if unit.kind != "word":
        return unit.kind in ("particle", "suffixed")
    p = _p_given_word(unit.core, vocab)
    return p is None or p < 0.6


def _unit_p_given(unit: _Unit, index: int, total: int, vocab: NameVocabulary) -> float:
    if unit.kind in ("particle", "suffixed"):
        return 0.02
    if unit.kind == "compound_given":
        return 0.99
    known = _p_given_word(unit.core, vocab)
    if known is not None:
        return min(max(known, 0.02), 0.98)
    # palavra desconhecida: só a posição ajuda (o 1º costuma ser nome próprio; o último, sobrenome)
    return 0.8 if index == 0 else 0.05 if index == total - 1 else 0.35


def split_full_name(full_name, vocab: NameVocabulary | None = None) -> NameSplit:
    """Divide o nome completo em Nome e Sobrenome, com a confiança da divisão e os motivos da dúvida."""
    vocab = vocab or get_vocabulary()
    words = sanitize_output_text(full_name, None).split()
    if not words:
        return NameSplit("", "", DOUBTFUL, ("nome vazio",))

    reasons: list[str] = []
    doubtful = False
    if any(ch.isdigit() for ch in "".join(words)):
        reasons.append("contém número")
        doubtful = True

    units = _build_units(words, vocab)
    total = len(units)

    if total == 1:
        reasons.append("só uma palavra: não dá para separar nome e sobrenome")
        return NameSplit(units[0].text, "", DOUBTFUL, tuple(reasons))

    if units[0].kind == "particle":
        reasons.append("começa com partícula")
        doubtful = True

    if total == 2:
        given_count = 1
        p_first = _unit_p_given(units[0], 0, total, vocab)
        p_last = _unit_p_given(units[1], 1, total, vocab)
        if p_first < 0.5 or p_last > 0.5:
            reasons.append("sem sobrenome reconhecido" if p_last > 0.5 else "sobrenome antes do nome?")
            doubtful = True
    else:
        first_forced_surname = next((i for i, u in enumerate(units) if u.kind in ("particle", "suffixed")), total)
        max_given = min(total - 1, first_forced_surname, max(_PRIOR_GIVEN_COUNT))
        if max_given < 1:
            given_count = 1
            reasons.append("não deu para achar o fim do nome")
            doubtful = True
        else:
            probabilities = [_unit_p_given(u, i, total, vocab) for i, u in enumerate(units)]
            scores = {}
            for candidate in range(1, max_given + 1):
                score = math.log(_PRIOR_GIVEN_COUNT[candidate])
                score += sum(math.log(p) for p in probabilities[:candidate])
                score += sum(math.log(1 - p) for p in probabilities[candidate:])
                scores[candidate] = score
            best_score = max(scores.values())
            weights = {g: math.exp(s - best_score) for g, s in scores.items()}
            given_count = max(weights, key=lambda g: weights[g])
            if weights[given_count] / sum(weights.values()) < _CONFIDENCE_THRESHOLD:
                reasons.append("a divisão entre nome e sobrenome é ambígua")
                doubtful = True

    if units[-1].kind == "word" and units[-1].text in SUFFIXES:
        # "Carlos Junior", "Maria Clara Neto": o sufixo sobrou sem sobrenome antes dele. Pode ser o sobrenome, pode ser
        # parte do nome; só quem conhece a pessoa sabe.
        reasons.append(f"'{units[-1].text}' é sufixo, mas não há sobrenome antes dele")
        doubtful = True

    if total > _MAX_UNITS_CONFIDENT:
        reasons.append(f"nome muito longo ({total} partes)")
        doubtful = True

    nome = " ".join(u.text for u in units[:given_count])
    sobrenome = " ".join(u.text for u in units[given_count:])

    # Motivos informativos (aparecem na conferência): o que a regra decidiu.
    if any(u.kind == "compound_given" for u in units):
        reasons.append("nome composto com partícula")
    elif given_count >= 2:
        reasons.append("nome composto")
    if any(u.kind == "particle" for u in units[given_count:]):
        reasons.append("partícula no sobrenome")
    if any(u.kind == "suffixed" for u in units[given_count:]) or any(w in SUFFIXES for w in words):
        reasons.append("sufixo no sobrenome")
    unknown = [u.core for u in units if u.kind == "word" and _p_given_word(u.core, vocab) is None]
    if unknown and (doubtful or given_count != 1):
        reasons.append("palavra fora do vocabulário: " + ", ".join(unknown))
    ambiguous = [
        u.core
        for u in units
        if u.kind == "word" and (p := _p_given_word(u.core, vocab)) is not None and 0.25 < p < 0.75
    ]
    if ambiguous:
        reasons.append("pode ser nome ou sobrenome: " + ", ".join(ambiguous))

    return NameSplit(nome, sobrenome, DOUBTFUL if doubtful else CONFIDENT, tuple(reasons))


# --- Tamanho ----------------------------------------------------------------------------------------------------
def suggest_abbreviation(nome: str, sobrenome: str, limit: int = MAX_FIELD_LEN) -> tuple[str, str]:
    """Proposta para quem passou de `limit`: nomes do meio viram inicial (o primeiro nome e o último sobrenome ficam
    inteiros). Devolve o que conseguir; se ainda não couber, quem confere edita à mão."""

    def shorten_given(text: str) -> str:
        parts = text.split()
        for index in range(len(parts) - 1, 0, -1):  # do último para o segundo: o 1º nome fica inteiro
            if len(" ".join(parts)) <= limit:
                break
            if parts[index] not in PARTICLES:
                parts[index] = parts[index][0]
        return " ".join(parts)

    def abbreviate_unit(unit: str) -> str:
        """ "DE ALBUQUERQUE" -> "DE A"; "SANTOS FILHO" -> "S FILHO" (a partícula e o sufixo ficam)."""
        words = unit.split()
        target = len(words) - 1
        if words[target] in SUFFIXES and target > 0:
            target -= 1
        words[target] = words[target][:1]
        return " ".join(words)

    def shorten_surname(text: str) -> str:
        units = [u.text for u in _build_units(text.split(), get_vocabulary(), suffix_min_units=1)]
        for index in range(len(units) - 1):  # da esquerda para a direita; o último sobrenome fica inteiro
            if len(" ".join(units)) <= limit:
                break
            units[index] = abbreviate_unit(units[index])
        return " ".join(units)

    return (
        shorten_given(nome) if len(nome) > limit else nome,
        shorten_surname(sobrenome) if len(sobrenome) > limit else sobrenome,
    )
