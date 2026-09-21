"""Inativação em cascata: usuário + estrutura direta do viajante + compactação de aprovadores.

Só a regra de negócio nova mora aqui; a mecânica de estruturas vem do `ApprovalService` e a busca
do `InactivationService`. Spec: docs/superpowers/specs/2026-09-21-inativacao-cascata-design.md
"""

import re
from dataclasses import dataclass
from typing import Any

import pandas as pd

from backend.core.config import settings
from backend.services.approval_service import ApprovalService
from backend.services.inactivation_service import InactivationService
from backend.shared.cpf_mask import mascarar_cpf
from backend.shared.cpf_utils import clean_cpf
from backend.shared.fingerprint import impressao_digital
from backend.shared.text_utils import upper_no_accents

ALERTA_SEM_CPF = (
    "Não foi possível mapear a Estrutura de Aprovação: Usuário encontrado no cadastro, mas não possui CPF registrado."
)

_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.IGNORECASE)


class InativacaoError(Exception):
    """Erro de negócio com código estável: o frontend decide pelo `code`, não pelo texto."""

    def __init__(self, code: str, message: str, status: int = 400, extra: dict[str, Any] | None = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.status = status
        self.extra = extra or {}


@dataclass(frozen=True)
class Analise:
    payload: dict[str, Any]  # o JSON de /analisar
    cpfs: frozenset[str]  # CPFs executáveis
    excluidas: frozenset[str]  # AprovacaoId das estruturas de viajante a excluir
    alvo_compactacao: frozenset[str]  # AprovacaoId a compactar (nunca as excluídas)
    orfas: frozenset[str]  # AprovacaoId que ficam sem nenhum aprovador
    df_estruturas: pd.DataFrame  # cópia com índice 0..n-1
    cols: dict[str, Any]


@dataclass(frozen=True)
class Execucao:
    ficha: pd.DataFrame  # ficha de inativação (Operacao=DELETE)
    estruturas: pd.DataFrame  # estruturas afetadas: DELETE nas excluídas, UPDATE nas alteradas
    resumo: dict[str, Any]


def _validar_cadastro(df: pd.DataFrame) -> None:
    cpf_col = InactivationService._detect_base_cols(df)[0]
    if not cpf_col:
        raise InativacaoError("BASE_SEM_COLUNA", "A base de cadastro não contém a coluna CPF.")


def _validar_estruturas(cols: dict[str, Any]) -> None:
    faltando = []
    if not cols.get("aprovacao_id"):
        faltando.append("AprovacaoId")
    if not cols.get("aprovacao_por"):
        faltando.append("AprovacaoPor")
    if not cols.get("approver_cols"):
        faltando.append("LoginAprovador_1")
    if not cols.get("traveler_cpf_col"):
        faltando.append("CPF (do viajante)")
    if faltando:
        raise InativacaoError("BASE_SEM_COLUNA", "A base de estruturas não contém: " + ", ".join(faltando) + ".")


def _classificar(itens: list[str]) -> tuple[set[str], set[str], dict[str, str]]:
    """(CPFs, e-mails em minúsculas, {nome normalizado: nome digitado}) digitados — espelha `search_matches`."""
    cpfs: set[str] = set()
    emails: set[str] = set()
    nomes: dict[str, str] = {}
    for item in itens:
        digits = re.sub(r"\D", "", item)
        if _EMAIL.match(item):
            emails.add(item.lower())
        elif len(digits) == 11:
            cpfs.add(digits)
        else:
            norm = upper_no_accents(item).strip()
            if len(norm.split()) >= 2:
                nomes[norm] = item
    return cpfs, emails, nomes


def _usuario(registro: dict[str, Any], situacao: str, alerta: str | None = None) -> dict[str, Any]:
    cpf = clean_cpf(registro.get("cpf"))
    return {
        "cpf": cpf or None,
        "cpfMascarado": mascarar_cpf(cpf),
        "nome": str(registro.get("nome", "")),
        "email": str(registro.get("email", "")),
        "situacao": situacao,
        "alerta": alerta,
        "estruturasViajante": [],
        "comoAprovador": [],
        "candidatos": [],
    }


def _nomes_homonimos(df_cadastro: pd.DataFrame, nomes_digitados: set[str]) -> set[str]:
    """Nomes digitados (normalizados) que aparecem em mais de uma linha do cadastro.

    Vem do cadastro, não da busca: `search_matches` omite da busca por nome quem já veio por CPF, então
    o grupo do nome pode ter 1 linha mesmo havendo homônimo.
    """
    nome_col = InactivationService._detect_base_cols(df_cadastro)[1]
    if not nome_col or not nomes_digitados:
        return set()
    contagem = df_cadastro[nome_col].map(lambda v: upper_no_accents(str(v)).strip()).value_counts()
    return {n for n in nomes_digitados if contagem.get(n, 0) > 1}


def _resolver_usuarios(
    busca: dict[str, Any], itens: list[str], escolhidos: set[str], df_cadastro: pd.DataFrame
) -> tuple[list[dict[str, Any]], set[str]]:
    """Transforma o resultado de `search_matches` em usuários com situação; devolve também os CPFs executáveis."""
    digitados_cpf, digitados_email, digitados_nome = _classificar(itens)
    homonimos = _nomes_homonimos(df_cadastro, set(digitados_nome))
    diretos: list[dict[str, Any]] = []
    por_nome: dict[str, list[dict[str, Any]]] = {}
    for r in (r for r in busca["items"] if r.get("found")):
        cpf = clean_cpf(r.get("cpf"))
        email = str(r.get("email", "")).strip().lower()
        norm = upper_no_accents(str(r.get("nome", ""))).strip()
        if (cpf and cpf in digitados_cpf) or (email and email in digitados_email):
            diretos.append(r)
        elif norm in digitados_nome:
            por_nome.setdefault(norm, []).append(r)
        else:
            diretos.append(r)

    usuarios: list[dict[str, Any]] = []
    cpfs: set[str] = set()
    vistos: set[str] = set()

    def adicionar(r: dict[str, Any]) -> None:
        cpf = clean_cpf(r.get("cpf"))
        if len(cpf) != 11:
            usuarios.append(_usuario(r, "SEM_CPF", ALERTA_SEM_CPF))
            return
        if cpf in vistos:
            return
        vistos.add(cpf)
        status = upper_no_accents(str(r.get("status_atual", ""))).strip()
        if status and status != "ATIVO":
            usuarios.append(_usuario(r, "JA_INATIVO", f"Usuário já consta como {status} no cadastro."))
            return
        usuarios.append(_usuario(r, "EXECUTAVEL"))
        cpfs.add(cpf)

    for r in diretos:
        adicionar(r)
    for norm, grupo in por_nome.items():
        if len(grupo) == 1 and norm not in homonimos:
            adicionar(grupo[0])
            continue
        escolhidos_do_grupo = [r for r in grupo if clean_cpf(r.get("cpf")) in escolhidos]
        if escolhidos_do_grupo:
            for r in escolhidos_do_grupo:
                adicionar(r)
            continue
        pendente = _usuario({"nome": digitados_nome[norm]}, "PENDENTE_SELECAO")
        pendente["candidatos"] = [
            {
                "cpf": clean_cpf(r.get("cpf")),
                "cpfMascarado": mascarar_cpf(r.get("cpf")),
                "nome": str(r.get("nome", "")),
                "email": str(r.get("email", "")),
            }
            for r in grupo
        ]
        usuarios.append(pendente)
    for r in (r for r in busca["items"] if not r.get("found")):
        usuarios.append(_usuario(r, "NAO_LOCALIZADO"))
    return usuarios, cpfs


class InactivationCascadeService:
    @staticmethod
    def analisar(
        df_cadastro: pd.DataFrame,
        df_estruturas: pd.DataFrame,
        itens: list[str],
        selecionados: list[str] | None = None,
    ) -> Analise:
        """Diagnóstico de impacto. Puro: não altera as entradas nem grava nada."""
        lista = [str(i).strip() for i in (itens or []) if str(i).strip()]
        if not lista:
            raise InativacaoError("LISTA_VAZIA", "Informe ao menos um CPF, nome completo ou e-mail.")
        if len(lista) > settings.MAX_INATIVACAO_ITENS:
            raise InativacaoError(
                "LISTA_GRANDE", f"A lista tem {len(lista)} itens; o máximo é {settings.MAX_INATIVACAO_ITENS}."
            )
        _validar_cadastro(df_cadastro)
        df_est = df_estruturas.copy().reset_index(drop=True)
        cols = ApprovalService.detect_approval_columns(df_est)
        _validar_estruturas(cols)

        escolhidos = {c for c in (clean_cpf(x) for x in (selecionados or [])) if c}
        busca = InactivationService.search_matches(df_cadastro, lista)
        usuarios, cpfs = _resolver_usuarios(busca, lista, escolhidos, df_cadastro)

        viajante = ApprovalService.find_traveler_structures(df_est, cpfs, cols)
        excluidas: set[str] = set().union(*viajante.values())
        como_aprovador = ApprovalService.find_approver_structures(df_est, cpfs, cols)
        alvo = {aid for mapa in como_aprovador.values() for aid in mapa} - excluidas
        orfas_info = ApprovalService.structures_left_without_approvers(df_est, cpfs, cols, alvo, True)
        orfas = {o["aprovacaoId"] for o in orfas_info}

        for usuario in usuarios:
            if usuario["situacao"] != "EXECUTAVEL":
                continue
            cpf = usuario["cpf"]
            usuario["estruturasViajante"] = sorted(viajante.get(cpf, set()))
            usuario["comoAprovador"] = [
                {
                    "aprovacaoId": aid,
                    "posicoes": info["posicoes"],
                    "segundoNivel": info["segundoNivel"],
                    "acao": "ORFA" if aid in orfas else "COMPACTACAO",
                }
                for aid, info in sorted(como_aprovador.get(cpf, {}).items())
                if aid not in excluidas
            ]

        compactadas = [
            {"cpf": cpf, "id": aid, "posicoes": info["posicoes"], "segundoNivel": info["segundoNivel"]}
            for cpf, mapa in como_aprovador.items()
            for aid, info in mapa.items()
            if aid not in excluidas
        ]
        digital = impressao_digital({"cpfs": cpfs, "excluidas": excluidas, "compactadas": compactadas, "orfas": orfas})
        payload = {
            "usuarios": usuarios,
            "resumo": {
                "executaveis": len(cpfs),
                "estruturasExcluidas": len(excluidas),
                "estruturasCompactadas": len(alvo - orfas),
                "estruturasOrfas": len(orfas),
                "duplicados": list(busca.get("duplicates", [])),
            },
            "impressaoDigital": digital,
        }
        return Analise(
            payload=payload,
            cpfs=frozenset(cpfs),
            excluidas=frozenset(excluidas),
            alvo_compactacao=frozenset(alvo),
            orfas=frozenset(orfas),
            df_estruturas=df_est,
            cols=cols,
        )

    @staticmethod
    def executar(
        df_cadastro: pd.DataFrame,
        df_estruturas: pd.DataFrame,
        cpfs: list[str],
        impressao_recebida: str,
        ignore_orphan_warning: bool = False,
    ) -> Execucao:
        """Recalcula a análise e, se ela bate com a que o operador viu, gera a ficha e as estruturas.

        Nunca confia no diagnóstico enviado pelo navegador: só na impressão digital dele.
        """
        lista = sorted({c for c in (clean_cpf(x) for x in (cpfs or [])) if c})
        if not lista:
            raise InativacaoError("NADA_A_EXECUTAR", "Nenhum usuário foi informado para inativar.")
        analise = InactivationCascadeService.analisar(df_cadastro, df_estruturas, lista, selecionados=lista)
        if not analise.cpfs:
            raise InativacaoError(
                "NADA_A_EXECUTAR",
                "Nenhum dos usuários informados pode ser inativado (não localizado, sem CPF ou já inativo).",
            )
        if analise.payload["impressaoDigital"] != impressao_recebida:
            raise InativacaoError(
                "ANALISE_DIVERGENTE",
                "A análise mudou desde a última conferência. Analise novamente antes de executar.",
                status=409,
            )
        if analise.orfas and not ignore_orphan_warning:
            raise InativacaoError(
                "ORFAS_SEM_CONFIRMACAO",
                f"{len(analise.orfas)} estrutura(s) ficará(ão) sem nenhum aprovador. Confirme para continuar.",
                extra={"estruturasOrfas": sorted(analise.orfas)},
            )

        ficha, _stats = InactivationService.process_from_dataframes(
            df_cadastro, pd.DataFrame({"CPF": sorted(analise.cpfs)})
        )
        if ficha.empty:
            raise InativacaoError(
                "NADA_A_EXECUTAR", "A ficha de inativação saiu vazia: nenhum usuário ATIVO correspondeu."
            )

        df_est, cols = analise.df_estruturas, analise.cols
        atualizado, stats = ApprovalService.remove_cpfs_and_compact(
            df_est, set(analise.cpfs), cols, set(analise.alvo_compactacao), True
        )
        id_col = cols["aprovacao_id"]
        compactadas = atualizado[atualizado[id_col].astype(str).str.strip().isin(analise.alvo_compactacao)].copy()
        compactadas["Operacao"] = ""
        compactadas.loc[compactadas.index.isin(stats["changed_indices"]), "Operacao"] = "UPDATE"
        excluidas = ApprovalService.delete_structures(df_est, set(analise.excluidas), cols)
        partes = [parte for parte in (excluidas, compactadas) if not parte.empty]
        estruturas = pd.concat(partes, ignore_index=True) if partes else compactadas
        estruturas = estruturas[["Operacao", *[c for c in estruturas.columns if c != "Operacao"]]]

        resumo = {
            "usuariosInativados": len(ficha),
            "estruturasExcluidas": len(analise.excluidas),
            "estruturasCompactadas": len(analise.alvo_compactacao - analise.orfas),
            "estruturasOrfas": len(analise.orfas),
            "linhasEstruturas": len(estruturas),
        }
        return Execucao(ficha=ficha, estruturas=estruturas, resumo=resumo)
