# Inativação em cascata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A aba Inativação passa a analisar o impacto (sem efeito colateral) e executar em cascata: inativa usuários, exclui a estrutura direta do viajante e compacta aprovadores, entregando um ZIP com 2 planilhas.

**Architecture:** Fluxo sem estado em duas rotas (`/api/inativacao/analisar` e `/executar`); o `/executar` recalcula a análise e compara uma impressão digital. A mecânica de estruturas é estendida no `ApprovalService` (passada única sobre a base); um orquestrador fino (`InactivationCascadeService`) liga busca, estruturas e ficha `DELETE`.

**Tech Stack:** Python 3.10+, Flask 3.1, pandas 2.3, openpyxl, pytest; React 19 + TypeScript, Vitest, Playwright.

**Spec:** `docs/superpowers/specs/2026-09-21-inativacao-cascata-design.md`

## Global Constraints

- CPF na auditoria sempre mascarado (`***.456.789-**`); nunca nome, e-mail nem CPF completo em `details`.
- Análise nunca altera os DataFrames de entrada; execução é determinística (mesmas entradas, mesma saída).
- Lista de até **500** itens (`MAX_INATIVACAO_ITENS`); teto de upload das rotas novas **32 MB** (`INATIVACAO_MAX_CONTENT_LENGTH`), aplicado por requisição.
- Erros com `code` estável e mensagem em português: `BASE_AUSENTE`, `ARQUIVO_INVALIDO`, `BASE_SEM_COLUNA`, `LISTA_VAZIA`, `LISTA_GRANDE`, `NADA_A_EXECUTAR`, `ORFAS_SEM_CONFIRMACAO`(400), `ANALISE_DIVERGENTE`(409), `ARQUIVO_GRANDE`(413), `ERRO_INTERNO`(500).
- Usuário sem CPF é **não executável**; alerta exato: `Não foi possível mapear a Estrutura de Aprovação: Usuário encontrado no cadastro, mas não possui CPF registrado.`
- Estrutura do viajante = linhas `AprovacaoPor=VIAJANTE` cuja coluna **CPF** da base de estruturas é o CPF do usuário; sai com `Operacao=DELETE`. Compactadas saem com `UPDATE`. O 2º nível é sempre removido (`remove_second_level=True`).
- `estruturas_atualizadas.xlsx` está **sempre** no ZIP (só cabeçalho quando não há mudança).
- Backend: ruff (`line-length = 120`, regras `E,F,I,UP,B`) e mypy limpos; rode os testes da raiz do repositório com `python -m pytest`.
- Todo commit termina com `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>`.
- Execute em uma branch/worktree própria a partir da `main` (`feat/inativacao-cascata`, via `superpowers:using-git-worktrees`); **não** dê push sem o usuário pedir.

## File Structure

**Criar (backend):**
- `backend/shared/cpf_mask.py` — `mascarar_cpf`
- `backend/shared/fingerprint.py` — `impressao_digital`
- `backend/services/inactivation_cascade_service.py` — `InativacaoError`, `Analise`, `Execucao`, `InactivationCascadeService`
- `backend/tests/_cascade_fixtures.py` — builders de planilhas dos testes
- `backend/tests/test_cpf_mask.py`, `test_fingerprint.py`, `test_approval_multi_cpf.py`, `test_inactivation_cascade_service.py`, `test_inactivation_executar.py`, `test_inativacao_analisar_api.py`, `test_inativacao_executar_api.py`

**Modificar (backend):** `backend/services/approval_service.py`, `backend/services/inactivation_service.py`, `backend/services/export_service.py`, `backend/api/inativacao.py` (reescrito), `backend/core/config.py`.

**Remover:** `backend/tests/test_inativacao_api.py`, `test_inativacao_buscar_api.py`; testes de rota em `test_inativacao_generation_api.py`, `test_dead_params_removed.py`, `test_upload_dir.py`.

**Criar (frontend):** `frontend-react/src/lib/inativacaoApi.ts` (+ `.test.ts`), `tabs/InativacaoTab/wizard.ts` (+ `.test.ts`), `tabs/InativacaoTab/ImpactoCard.tsx` (+ `.test.tsx`).
**Modificar (frontend):** `lib/api.ts`, `tabs/InativacaoTab/useInativacao.ts` (+ teste), `tabs/InativacaoTab/index.tsx`.
**Outros:** `tests/e2e/inativacao-flow.spec.js`, `REGRAS_APROVACAO_INATIVACAO.md`, `ARQUITETURA_MODERNIZADA.md`, `README.md`.

---

### Task 1: Máscara de CPF e impressão digital

**Files:**
- Create: `backend/shared/cpf_mask.py`, `backend/shared/fingerprint.py`
- Test: `backend/tests/test_cpf_mask.py`, `backend/tests/test_fingerprint.py`

**Interfaces:**
- Produces: `mascarar_cpf(value) -> str` (`"***.456.789-**"` ou `""`); `impressao_digital(payload: dict) -> str` (SHA-256 hex, 64 caracteres, independente da ordem de chaves/listas/conjuntos).

- [ ] **Step 1: Write the failing tests**

`backend/tests/test_cpf_mask.py`:

```python
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
```

`backend/tests/test_fingerprint.py`:

```python
from backend.shared.fingerprint import impressao_digital


def test_mesma_informacao_em_outra_ordem_gera_a_mesma_impressao():
    a = {"cpfs": ["2", "1"], "excluidas": ["A"], "compactadas": [{"id": "B", "posicoes": [3, 1]}]}
    b = {"compactadas": [{"posicoes": [1, 3], "id": "B"}], "excluidas": ["A"], "cpfs": ["1", "2"]}
    assert impressao_digital(a) == impressao_digital(b)


def test_impacto_diferente_muda_a_impressao():
    assert impressao_digital({"cpfs": ["1"], "orfas": []}) != impressao_digital({"cpfs": ["1"], "orfas": ["A"]})


def test_aceita_conjuntos_e_devolve_sha256_hex():
    digest = impressao_digital({"cpfs": {"1", "2"}})
    assert len(digest) == 64
    assert int(digest, 16) >= 0
    assert digest == impressao_digital({"cpfs": frozenset({"2", "1"})})
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest backend/tests/test_cpf_mask.py backend/tests/test_fingerprint.py -q`
Expected: FAIL — `ModuleNotFoundError: backend.shared.cpf_mask`.

- [ ] **Step 3: Write minimal implementation**

`backend/shared/cpf_mask.py`:

```python
from backend.shared.cpf_utils import clean_cpf


def mascarar_cpf(value) -> str:
    """``12345678909`` -> ``***.456.789-**`` (só os 6 dígitos do meio); ``""`` se não houver CPF de 11 dígitos."""
    digits = clean_cpf(value)
    if len(digits) != 11:
        return ""
    return f"***.{digits[3:6]}.{digits[6:9]}-**"
```

`backend/shared/fingerprint.py`:

```python
import hashlib
import json
from typing import Any


def _canonico(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): _canonico(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set, frozenset)):
        items = [_canonico(v) for v in value]
        return sorted(items, key=lambda v: json.dumps(v, sort_keys=True, ensure_ascii=False))
    return value


def impressao_digital(payload: dict[str, Any]) -> str:
    """SHA-256 do JSON canônico (chaves, listas e conjuntos ordenados): identifica o CONTEÚDO do diagnóstico."""
    canonico = json.dumps(_canonico(payload), sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(canonico.encode("utf-8")).hexdigest()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest backend/tests/test_cpf_mask.py backend/tests/test_fingerprint.py -q`
Expected: PASS (7 testes).

- [ ] **Step 5: Commit**

```bash
git add backend/shared/cpf_mask.py backend/shared/fingerprint.py backend/tests/test_cpf_mask.py backend/tests/test_fingerprint.py
git commit -m "feat(backend): máscara de CPF e impressão digital do diagnóstico"
```

---

### Task 2: `ApprovalService` — passada única para vários CPFs

**Files:**
- Modify: `backend/services/approval_service.py` (helpers no topo; `detect_approval_columns`; `remove_cpf_and_compact`; `_check_structures_without_approvers`; `build_preview_for_cpf`; novos métodos)
- Test: `backend/tests/test_approval_multi_cpf.py`

**Interfaces:**
- Produces (todos em `ApprovalService`, estáticos, sem I/O e sem mutar a entrada):
  - `detect_approval_columns(df)` passa a devolver também `"traveler_cpf_col"` (coluna `CPF`/`CPFViajante`/`CPFDoViajante`, ou `None`).
  - `find_traveler_structures(df, cpfs: set[str], cols) -> dict[str, set[str]]` — `{cpf: {AprovacaoId}}`; `ValueError("Base de estruturas não contém a coluna CPF do viajante.")` se faltar a coluna.
  - `find_approver_structures(df, cpfs, cols) -> dict[str, dict[str, dict]]` — `{cpf: {aid: {"posicoes": [int], "segundoNivel": bool}}}`.
  - `remove_cpfs_and_compact(df, cpfs: set[str], cols, target_ids: set[str], remove_second_level: bool) -> tuple[DataFrame, dict]` (mesmas chaves de stats que `remove_cpf_and_compact`).
  - `structures_left_without_approvers(df, cpfs, cols, target_ids, remove_second_level) -> list[dict]` (`aprovacaoId`, `aprovacaoPor`, `valor`, `contexto`).
  - `delete_structures(df, ids: set[str], cols) -> DataFrame` — linhas dessas estruturas com `Operacao="DELETE"`.
- `remove_cpf_and_compact(...)` mantém a assinatura e passa a delegar a `remove_cpfs_and_compact` com `{cpf_digits}`.

- [ ] **Step 1: Write the failing tests**

`backend/tests/test_approval_multi_cpf.py`:

```python
import pandas as pd
import pytest
from _helpers import format_cpf, valid_cpf

from backend.services.approval_service import ApprovalService

A, B, C, D = (valid_cpf(i) for i in (1, 2, 3, 4))


def _df(*rows):
    return pd.DataFrame(list(rows)).fillna("").astype(str)


def _cols(df):
    return ApprovalService.detect_approval_columns(df)


def _linha(aid, *aprovadores, por="CCEMPRESA", cpf="", segundo=""):
    linha = {"AprovacaoId": aid, "AprovacaoPor": por, "CPF": cpf}
    for i, login in enumerate(aprovadores, 1):
        linha[f"LoginAprovador_{i}"] = login
    if segundo:
        linha["LoginAprovador_SEGUNDO_NIVEL"] = segundo
    return linha


def test_detecta_coluna_cpf_do_viajante():
    assert _cols(_df(_linha("S1", A)))["traveler_cpf_col"] == "CPF"


def test_find_traveler_structures_so_conta_viajante_e_aceita_pontuacao():
    df = _df(
        _linha("S1", B, por="VIAJANTE", cpf=format_cpf(A)),
        _linha("S2", B, por="VIAJANTE", cpf=B),
        _linha("S3", B, por="CCEMPRESA", cpf=A),
    )
    assert ApprovalService.find_traveler_structures(df, {A, B}, _cols(df)) == {A: {"S1"}, B: {"S2"}}


def test_find_traveler_structures_sem_coluna_cpf_levanta_erro():
    df = _df(_linha("S1", A)).drop(columns=["CPF"])
    with pytest.raises(ValueError, match="CPF do viajante"):
        ApprovalService.find_traveler_structures(df, {A}, _cols(df))


def test_find_approver_structures_devolve_posicoes_e_segundo_nivel():
    df = _df(_linha("S1", A, B), _linha("S2", C, A, segundo=A))
    achados = ApprovalService.find_approver_structures(df, {A, B}, _cols(df))
    assert achados[A] == {
        "S1": {"posicoes": [1], "segundoNivel": False},
        "S2": {"posicoes": [2], "segundoNivel": True},
    }
    assert achados[B] == {"S1": {"posicoes": [2], "segundoNivel": False}}


def test_remove_varios_cpfs_compacta_uma_vez():
    df = _df(_linha("S1", A, B, C, D))
    out, stats = ApprovalService.remove_cpfs_and_compact(df, {A, C}, _cols(df), {"S1"}, True)
    assert out.loc[0, ["LoginAprovador_1", "LoginAprovador_2", "LoginAprovador_3", "LoginAprovador_4"]].tolist() == [
        B,
        D,
        "",
        "",
    ]
    assert stats["structures_updated"] == 1
    assert stats["occurrences_removed"] == 2


def test_remove_varios_cpfs_equivale_a_remocoes_sequenciais():
    df = _df(_linha("S1", A, B, C), _linha("S2", B, A), _linha("S3", C), _linha("S4", D, A, segundo=B))
    cols = _cols(df)
    ids = {"S1", "S2", "S3", "S4"}
    de_uma_vez, _ = ApprovalService.remove_cpfs_and_compact(df, {A, B}, cols, ids, True)
    passo, _ = ApprovalService.remove_cpf_and_compact(df, A, cols, ids, True)
    em_sequencia, _ = ApprovalService.remove_cpf_and_compact(passo, B, cols, ids, True)
    assert de_uma_vez.equals(em_sequencia)


def test_promove_segundo_nivel_quando_o_primeiro_esvazia():
    df = _df(_linha("S1", A, segundo=C))
    out, stats = ApprovalService.remove_cpfs_and_compact(df, {A}, _cols(df), {"S1"}, True)
    assert out.loc[0, "LoginAprovador_1"] == C
    assert out.loc[0, "LoginAprovador_SEGUNDO_NIVEL"] == ""
    assert stats["promotions"] == 1


def test_segundo_nivel_de_outro_cpf_da_lista_nao_e_promovido():
    df = _df(_linha("S1", A, segundo=B))
    out, stats = ApprovalService.remove_cpfs_and_compact(df, {A, B}, _cols(df), {"S1"}, True)
    assert out.loc[0, "LoginAprovador_1"] == ""
    assert out.loc[0, "LoginAprovador_SEGUNDO_NIVEL"] == ""
    assert stats["promotions"] == 0


def test_structures_left_without_approvers_considera_todos_os_cpfs():
    df = _df(_linha("S1", A, B), _linha("S2", A, C))
    orfas = ApprovalService.structures_left_without_approvers(df, {A, B}, _cols(df), {"S1", "S2"}, True)
    assert [o["aprovacaoId"] for o in orfas] == ["S1"]


def test_delete_structures_marca_operacao_sem_mexer_na_entrada():
    df = _df(_linha("S1", A, por="VIAJANTE", cpf=A), _linha("S2", B))
    saida = ApprovalService.delete_structures(df, {"S1"}, _cols(df))
    assert saida["AprovacaoId"].tolist() == ["S1"]
    assert saida["Operacao"].tolist() == ["DELETE"]
    assert "Operacao" not in df.columns


def test_nao_altera_a_entrada():
    df = _df(_linha("S1", A, B), _linha("S2", C, segundo=A))
    antes = df.copy()
    ApprovalService.remove_cpfs_and_compact(df, {A}, _cols(df), {"S1", "S2"}, True)
    assert df.equals(antes)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest backend/tests/test_approval_multi_cpf.py -q`
Expected: FAIL — `KeyError: 'traveler_cpf_col'` / `AttributeError: ... has no attribute 'find_traveler_structures'`.

- [ ] **Step 3: Add the helpers and detect the CPF column**

Em `backend/services/approval_service.py`, logo depois dos imports (antes de `_get_or_create_structure`), adicione:

```python
def _digits_matrix(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """CPF (11 dígitos) de cada célula das colunas dadas; "" onde a célula está vazia.

    Cada valor distinto é normalizado uma única vez: as bases repetem muito o mesmo login e
    `limpar_cpf_raw` (regex) é o custo dominante.
    """
    cache: dict[str, str] = {}
    out: dict[str, pd.Series] = {}
    for col in columns:
        raw = df[col].astype(str).str.strip()
        for value in raw.unique():
            if value and value not in cache:
                cache[value] = limpar_cpf_raw(value)
        out[col] = raw.map(cache).fillna("")
    return pd.DataFrame(out, index=df.index)


def _orphan_rows_mask(
    df_base: pd.DataFrame,
    cpfs: set[str],
    cols: dict[str, Any],
    target_ids: set[str],
    remove_second_level: bool,
) -> pd.Series:
    """Linhas das estruturas-alvo que terminam SEM nenhum aprovador (vetorizado)."""
    approver_cols: list[str] = cols["approver_cols"]
    login_segundo_col = cols.get("login_segundo")
    wanted = list(cpfs)
    in_target = df_base[cols["aprovacao_id"]].astype(str).str.strip().isin(target_ids)
    raw = df_base[approver_cols].astype(str).apply(lambda col: col.str.strip())
    remaining = ((raw != "") & ~_digits_matrix(df_base, approver_cols).isin(wanted)).any(axis=1)
    has_second = pd.Series(False, index=df_base.index)
    if login_segundo_col and login_segundo_col in df_base.columns:
        second_raw = df_base[login_segundo_col].astype(str).str.strip()
        removed = pd.Series(False, index=df_base.index)
        if remove_second_level:
            removed = _digits_matrix(df_base, [login_segundo_col])[login_segundo_col].isin(wanted)
        has_second = (second_raw != "") & ~removed
    return in_target & ~remaining & ~has_second
```

Em `detect_approval_columns`, depois de `traveler_name_col = pick(...)` acrescente:

```python
        traveler_cpf_col = pick("CPFViajante", "CPFDoViajante", "CPF")
```

e no dicionário retornado, depois de `"traveler_name_col": traveler_name_col,`:

```python
            "traveler_cpf_col": traveler_cpf_col,
```

- [ ] **Step 4: Generalize `_check_structures_without_approvers` (CPFs em conjunto)**

Quatro edições em `_check_structures_without_approvers` (as trocas são únicas no arquivo):

1. Assinatura: `    cpf_digits: str,\n    cols: dict[str, Any],\n    target_ids: set[str],\n    remove_second_level: bool,\n) -> list[dict[str, Any]]:\n    """Verifica quais estruturas ficarão sem aprovadores após a remoção do CPF.` → trocar `cpf_digits: str,` por `cpfs: set[str],` e o docstring por `"""Verifica quais estruturas ficarão sem aprovadores após a remoção dos CPFs.`
2. `if limpar_cpf_raw(raw_login) == cpf_digits:\n                continue` → `if limpar_cpf_raw(raw_login) in cpfs:\n                continue`
3. `if remove_second_level and limpar_cpf_raw(raw_second) == cpf_digits:` → `if remove_second_level and limpar_cpf_raw(raw_second) in cpfs:`
4. `structures_without_approvers: list[dict[str, Any]] = []\n\n    for _idx, row in df_base.iterrows():` → 

```python
    structures_without_approvers: list[dict[str, Any]] = []

    # O laço só monta o registro das linhas que de fato ficam vazias (o filtro é vetorizado).
    orphan_mask = _orphan_rows_mask(df_base, cpfs, cols, target_ids, remove_second_level)
    for _idx, row in df_base[orphan_mask].iterrows():
```

No chamador, dentro de `build_preview_for_cpf`, troque `                cpf_digits=cpf_digits,\n                cols=cols,\n                target_ids=affected_ids,` por `                cpfs={cpf_digits},\n                cols=cols,\n                target_ids=affected_ids,`.

- [ ] **Step 5: Add the new methods and delegate `remove_cpf_and_compact`**

Dentro de `class ApprovalService`, substitua **todo o corpo** de `remove_cpf_and_compact` (mantendo a assinatura e o docstring) por:

```python
        return ApprovalService.remove_cpfs_and_compact(df_base, {cpf_digits}, cols, target_ids, remove_second_level)
```

e acrescente, logo antes de `check_new_approver_duplicates`:

```python
    @staticmethod
    def remove_cpfs_and_compact(
        df_base: pd.DataFrame,
        cpfs: set[str],
        cols: dict[str, Any],
        target_ids: set[str],
        remove_second_level: bool,
    ) -> tuple[pd.DataFrame, dict[str, Any]]:
        """Remove TODOS os `cpfs` de LoginAprovador_1..100 (e, se pedido, do 2º nível) e compacta.

        Uma passada só: só as linhas das estruturas-alvo em que algum CPF aparece passam pelo laço.
        Não altera `df_base`.
        """
        approver_cols: list[str] = cols.get("approver_cols") or []
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not approver_cols or not cpfs:
            return df_base, {
                "structures_updated": 0,
                "occurrences_removed": 0,
                "changed_indices": set(),
                "promotions": 0,
            }

        df_out = df_base.copy()
        login_segundo_col = cols.get("login_segundo")
        has_segundo = bool(login_segundo_col and login_segundo_col in df_out.columns)
        segundo_master_col = cols.get("segundo_master")
        wanted = list(cpfs)

        in_target = df_out[aprov_id_col].astype(str).str.strip().isin(target_ids)
        has_main = _digits_matrix(df_out, approver_cols).isin(wanted).any(axis=1)
        if has_segundo:
            has_second = _digits_matrix(df_out, [login_segundo_col])[login_segundo_col].isin(wanted)
        else:
            has_second = pd.Series(False, index=df_out.index)
        candidates = df_out.index[(in_target & (has_main | has_second)).to_numpy()]

        structures_updated: set[str] = set()
        changed_indices: set[Any] = set()
        occurrences_removed = 0
        promotions = 0

        for idx in candidates:
            row = df_out.loc[idx]
            aprov_id = str(row.get(aprov_id_col, "")).strip()
            changed = False

            if bool(has_main.at[idx]):
                kept: list[str] = []
                for col in approver_cols:
                    value = str(row.get(col, ""))
                    digits = limpar_cpf_raw(value)
                    if digits and digits in cpfs:
                        occurrences_removed += 1
                        changed = True
                        continue
                    if value.strip():
                        kept.append(value)
                for pos, col in enumerate(approver_cols):
                    df_out.at[idx, col] = kept[pos] if pos < len(kept) else ""

            if remove_second_level and has_segundo and bool(has_second.at[idx]):
                df_out.at[idx, login_segundo_col] = ""
                occurrences_removed += 1
                changed = True

            # Se o 1º nível esvaziou e sobrou um 2º nível que NÃO está sendo removido, ele sobe.
            if has_segundo:
                remaining_main = [
                    str(df_out.at[idx, col]).strip() for col in approver_cols if str(df_out.at[idx, col]).strip()
                ]
                current_second = str(df_out.at[idx, login_segundo_col]).strip()
                if not remaining_main and current_second and limpar_cpf_raw(current_second) not in cpfs:
                    df_out.at[idx, approver_cols[0]] = current_second
                    df_out.at[idx, login_segundo_col] = ""
                    promotions += 1
                    changed = True

            if changed:
                structures_updated.add(aprov_id)
                changed_indices.add(idx)

        if segundo_master_col and segundo_master_col in df_out.columns:
            df_out[segundo_master_col] = df_out[segundo_master_col].astype(str).fillna("")

        return df_out, {
            "structures_updated": len(structures_updated),
            "occurrences_removed": int(occurrences_removed),
            "promotions": int(promotions),
            "changed_indices": changed_indices,
        }

    @staticmethod
    def structures_left_without_approvers(
        df_base: pd.DataFrame,
        cpfs: set[str],
        cols: dict[str, Any],
        target_ids: set[str],
        remove_second_level: bool,
    ) -> list[dict[str, Any]]:
        """Estruturas de `target_ids` que ficam SEM nenhum aprovador depois de remover todos os `cpfs`."""
        if not cpfs or not target_ids:
            return []
        return _check_structures_without_approvers(df_base, set(cpfs), cols, set(target_ids), remove_second_level)

    @staticmethod
    def find_traveler_structures(df_base: pd.DataFrame, cpfs: set[str], cols: dict[str, Any]) -> dict[str, set[str]]:
        """`{cpf: {AprovacaoId}}` das estruturas AprovacaoPor=VIAJANTE cujo CPF do viajante é o do usuário."""
        cpf_col = cols.get("traveler_cpf_col")
        if not cpf_col:
            raise ValueError("Base de estruturas não contém a coluna CPF do viajante.")
        found: dict[str, set[str]] = {cpf: set() for cpf in cpfs}
        aprov_id_col = cols.get("aprovacao_id")
        por_col = cols.get("aprovacao_por")
        if not aprov_id_col or not por_col or not cpfs:
            return found
        is_traveler = df_base[por_col].astype(str).str.strip().str.upper() == "VIAJANTE"
        digits = _digits_matrix(df_base, [cpf_col])[cpf_col]
        ids = df_base[aprov_id_col].astype(str).str.strip()
        hit = is_traveler & digits.isin(list(cpfs)) & (ids != "")
        for idx in df_base.index[hit.to_numpy()]:
            found[digits.at[idx]].add(ids.at[idx])
        return found

    @staticmethod
    def find_approver_structures(
        df_base: pd.DataFrame, cpfs: set[str], cols: dict[str, Any]
    ) -> dict[str, dict[str, dict[str, Any]]]:
        """`{cpf: {aprovacaoId: {"posicoes": [1, 3], "segundoNivel": False}}}` — onde cada CPF é aprovador."""
        result: dict[str, dict[str, dict[str, Any]]] = {cpf: {} for cpf in cpfs}
        approver_cols: list[str] = cols.get("approver_cols") or []
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not approver_cols or not cpfs:
            return result
        wanted = list(cpfs)
        ids = df_base[aprov_id_col].astype(str).str.strip()
        digits = _digits_matrix(df_base, approver_cols)
        for col in approver_cols:
            hit = digits[col].isin(wanted) & (ids != "")
            if not hit.any():
                continue
            match = re.search(r"(\d+)$", str(col))
            pos = int(match.group(1)) if match else 0
            for idx in df_base.index[hit.to_numpy()]:
                entry = result[digits.at[idx, col]].setdefault(ids.at[idx], {"posicoes": [], "segundoNivel": False})
                if pos not in entry["posicoes"]:
                    entry["posicoes"].append(pos)
        second_col = cols.get("login_segundo")
        if second_col and second_col in df_base.columns:
            second = _digits_matrix(df_base, [second_col])[second_col]
            hit = second.isin(wanted) & (ids != "")
            for idx in df_base.index[hit.to_numpy()]:
                entry = result[second.at[idx]].setdefault(ids.at[idx], {"posicoes": [], "segundoNivel": False})
                entry["segundoNivel"] = True
        for by_id in result.values():
            for entry in by_id.values():
                entry["posicoes"].sort()
        return result

    @staticmethod
    def delete_structures(df_base: pd.DataFrame, ids: set[str], cols: dict[str, Any]) -> pd.DataFrame:
        """Linhas das estruturas em `ids` com Operacao=DELETE. Não altera `df_base`."""
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not ids:
            return df_base.iloc[0:0].copy()
        rows = df_base[df_base[aprov_id_col].astype(str).str.strip().isin(ids)].copy()
        rows["Operacao"] = "DELETE"
        return rows
```

- [ ] **Step 6: Run the new tests and the whole approval suite (regression)**

Run: `python -m pytest backend/tests/test_approval_multi_cpf.py -q`
Expected: PASS (11 testes). Se `test_remove_varios_cpfs_equivale_a_remocoes_sequenciais` falhar, é divergência real entre a passada única e o comportamento antigo: corrija o método novo, **não** o teste.

Run: `python -m pytest backend/tests/test_aprovacao.py backend/tests/test_aprovacao_substituir.py -q`
Expected: PASS sem nenhuma alteração nesses arquivos. Se algum teste comparar o dicionário de `detect_approval_columns` por igualdade, acrescente só a chave `traveler_cpf_col` na expectativa.

- [ ] **Step 7: Commit**

```bash
git add backend/services/approval_service.py backend/tests/test_approval_multi_cpf.py
git commit -m "refactor(backend): ApprovalService remove vários CPFs numa passada e localiza estruturas de viajante"
```

### Task 3: Análise de impacto (`InactivationCascadeService.analisar`)

**Files:**
- Modify: `backend/core/config.py` (2 limites), `backend/services/inactivation_service.py` (`search_matches` restaura zero à esquerda)
- Create: `backend/services/inactivation_cascade_service.py`, `backend/tests/_cascade_fixtures.py`
- Test: `backend/tests/test_inactivation_cascade_service.py`, `backend/tests/test_inativacao_search.py` (1 teste novo)

**Interfaces:**
- Consumes: Task 1 (`mascarar_cpf`, `impressao_digital`), Task 2 (`ApprovalService.find_traveler_structures`, `find_approver_structures`, `structures_left_without_approvers`).
- Produces:
  - `InativacaoError(code: str, message: str, status: int = 400, extra: dict | None = None)` com atributos `code`, `message`, `status`, `extra`.
  - `ALERTA_SEM_CPF` (texto exato das Global Constraints).
  - `Analise` (dataclass imutável): `payload: dict` (o JSON de `/analisar`), `cpfs`, `excluidas`, `alvo_compactacao`, `orfas` (`frozenset[str]`), `df_estruturas: DataFrame`, `cols: dict`.
  - `InactivationCascadeService.analisar(df_cadastro, df_estruturas, itens: list[str], selecionados: list[str] | None = None) -> Analise`.
  - `payload = {"usuarios": [...], "resumo": {"executaveis", "estruturasExcluidas", "estruturasCompactadas", "estruturasOrfas", "duplicados"}, "impressaoDigital": str}`; cada usuário: `cpf | None`, `cpfMascarado`, `nome`, `email`, `situacao` (`EXECUTAVEL|SEM_CPF|JA_INATIVO|NAO_LOCALIZADO|PENDENTE_SELECAO`), `alerta`, `estruturasViajante: [str]`, `comoAprovador: [{aprovacaoId, posicoes, segundoNivel, acao: COMPACTACAO|ORFA}]`, `candidatos: [{cpf, cpfMascarado, nome, email}]`.

- [ ] **Step 1: Builders de teste compartilhados**

`backend/tests/_cascade_fixtures.py`:

```python
"""Builders de planilhas para os testes da inativação em cascata."""

import pandas as pd
from _helpers import valid_cpf

A, B, C, D = (valid_cpf(i) for i in (1, 2, 3, 4))

USR_A = (A, "Ana Souza", "ana@x.com", "ATIVO")
USR_B = (B, "Bruno Lima", "bruno@x.com", "ATIVO")
USR_C = (C, "Carla Dias", "carla@x.com", "ATIVO")


def cad(*usuarios):
    """Base de cadastro; cada usuário é `(cpf, nome, email, status)`."""
    return pd.DataFrame(
        [{"CPF": c, "NomeCompleto": n, "Email": e, "Status": s} for c, n, e, s in usuarios]
    ).astype(str)


def est(*linhas):
    """Base de estruturas a partir de linhas (dicts); células ausentes viram texto vazio."""
    return pd.DataFrame(list(linhas)).fillna("").astype(str)


def viajante(aid, cpf, *aprovadores, segundo=""):
    """Estrutura AprovacaoPor=VIAJANTE do viajante `cpf`, com os aprovadores em 1..n."""
    linha = {"AprovacaoId": aid, "AprovacaoPor": "VIAJANTE", "CPF": cpf, "NomeViajante": f"Viajante {aid}"}
    for i, login in enumerate(aprovadores, 1):
        linha[f"LoginAprovador_{i}"] = login
    if segundo:
        linha["LoginAprovador_SEGUNDO_NIVEL"] = segundo
    return linha
```

- [ ] **Step 2: Write the failing tests**

`backend/tests/test_inativacao_search.py` — acrescente ao final:

```python
def test_busca_por_cpf_restaura_zero_a_esquerda_da_base():
    import pandas as pd

    from backend.services.inactivation_service import InactivationService

    base = pd.DataFrame([{"CPF": "1234567890", "NomeCompleto": "Ana Souza", "Email": "a@x.com", "Status": "ATIVO"}])
    achados = InactivationService.search_matches(base, ["01234567890"])
    assert [i["found"] for i in achados["items"]] == [True]
    assert achados["items"][0]["cpf"] == "01234567890"
```

`backend/tests/test_inactivation_cascade_service.py`:

```python
import pytest
from _cascade_fixtures import A, B, C, D, USR_A, USR_B, USR_C, cad, est, viajante
from _helpers import valid_cpf

from backend.core.config import settings
from backend.services.inactivation_cascade_service import (
    ALERTA_SEM_CPF,
    InactivationCascadeService,
    InativacaoError,
)


def _analisar(cadastro, estruturas, itens, **kw):
    return InactivationCascadeService.analisar(cadastro, estruturas, itens, **kw)


def _usuario(analise, indice=0):
    return analise.payload["usuarios"][indice]


def test_estrutura_do_viajante_e_marcada_para_exclusao():
    a = _analisar(cad(USR_A), est(viajante("S1", A, B)), [A])
    u = _usuario(a)
    assert u["situacao"] == "EXECUTAVEL"
    assert u["estruturasViajante"] == ["S1"]
    assert a.payload["resumo"]["estruturasExcluidas"] == 1
    assert a.excluidas == {"S1"}


def test_aprovador_com_substitutos_vira_compactacao():
    a = _analisar(cad(USR_A), est(viajante("S1", C, B, A)), [A])
    u = _usuario(a)
    assert u["estruturasViajante"] == []
    assert u["comoAprovador"] == [{"aprovacaoId": "S1", "posicoes": [2], "segundoNivel": False, "acao": "COMPACTACAO"}]
    assert a.payload["resumo"]["estruturasCompactadas"] == 1


def test_aprovador_unico_vira_estrutura_orfa():
    a = _analisar(cad(USR_A), est(viajante("S1", C, A)), [A])
    assert _usuario(a)["comoAprovador"][0]["acao"] == "ORFA"
    assert a.payload["resumo"]["estruturasOrfas"] == 1
    assert a.orfas == {"S1"}


def test_dois_da_lista_que_aprovam_uma_estrutura_deixam_ela_orfa_para_ambos():
    a = _analisar(cad(USR_A, USR_B), est(viajante("S1", C, A, B)), [A, B])
    assert {u["cpf"]: u["comoAprovador"][0]["acao"] for u in a.payload["usuarios"]} == {A: "ORFA", B: "ORFA"}
    assert a.payload["resumo"]["estruturasOrfas"] == 1


def test_estrutura_excluida_nunca_conta_como_orfa():
    a = _analisar(cad(USR_A), est(viajante("S1", A, A)), [A])
    assert _usuario(a)["estruturasViajante"] == ["S1"]
    assert _usuario(a)["comoAprovador"] == []
    assert a.payload["resumo"]["estruturasOrfas"] == 0


def test_segundo_nivel_e_removido_e_conta_como_compactacao():
    a = _analisar(cad(USR_A), est(viajante("S1", C, B, segundo=A)), [A])
    assert _usuario(a)["comoAprovador"] == [
        {"aprovacaoId": "S1", "posicoes": [], "segundoNivel": True, "acao": "COMPACTACAO"}
    ]


def test_usuario_sem_cpf_nao_e_executavel_e_traz_o_alerta_exato():
    a = _analisar(cad(("", "Sem Cpf Silva", "s@x.com", "ATIVO")), est(viajante("S1", C, B)), ["Sem Cpf Silva"])
    u = _usuario(a)
    assert u["situacao"] == "SEM_CPF"
    assert u["alerta"] == ALERTA_SEM_CPF
    assert a.cpfs == frozenset()


def test_usuario_ja_inativo_e_nao_localizado():
    a = _analisar(cad((A, "Ana Souza", "ana@x.com", "INATIVO")), est(viajante("S1", C, B)), [A, valid_cpf(9)])
    assert [u["situacao"] for u in a.payload["usuarios"]] == ["JA_INATIVO", "NAO_LOCALIZADO"]
    assert a.cpfs == frozenset()


def test_homonimos_exigem_selecao_explicita():
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    estruturas = est(viajante("S9", C, D))
    pendente = _analisar(cadastro, estruturas, ["Joao Silva"])
    u = _usuario(pendente)
    assert u["situacao"] == "PENDENTE_SELECAO"
    assert {c["cpf"] for c in u["candidatos"]} == {A, B}
    assert pendente.cpfs == frozenset()

    escolhido = _analisar(cadastro, estruturas, ["Joao Silva"], selecionados=[A])
    assert [u["situacao"] for u in escolhido.payload["usuarios"]] == ["EXECUTAVEL"]
    assert escolhido.cpfs == {A}


def test_cpf_com_pontuacao_e_zero_a_esquerda_perdido():
    cadastro = cad(("1234567890", "Ana Souza", "ana@x.com", "ATIVO"))
    a = _analisar(cadastro, est(viajante("S1", C, "012.345.678-90", B)), ["01234567890"])
    assert _usuario(a)["cpf"] == "01234567890"
    assert _usuario(a)["comoAprovador"][0]["aprovacaoId"] == "S1"


def test_cpf_repetido_na_lista_entra_uma_vez():
    a = _analisar(cad(USR_A), est(viajante("S1", C, B)), [A, A])
    assert len(a.payload["usuarios"]) == 1
    assert a.payload["resumo"]["duplicados"] == [A]


def test_nao_altera_as_entradas():
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, A))
    antes_c, antes_e = cadastro.copy(), estruturas.copy()
    _analisar(cadastro, estruturas, [A])
    assert cadastro.equals(antes_c)
    assert estruturas.equals(antes_e)


def test_impressao_digital_estavel_e_sensivel_ao_impacto():
    cadastro = cad(USR_A)
    base = est(viajante("S1", C, A, B))
    primeira = _analisar(cadastro, base, [A]).payload["impressaoDigital"]
    assert primeira == _analisar(cadastro, base, [A]).payload["impressaoDigital"]
    outra = est(viajante("S1", C, A))  # agora A é o único aprovador
    assert primeira != _analisar(cadastro, outra, [A]).payload["impressaoDigital"]


def test_lista_vazia():
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), est(viajante("S1", C, B)), ["  "])
    assert erro.value.code == "LISTA_VAZIA"


def test_lista_grande(monkeypatch):
    monkeypatch.setattr(settings, "MAX_INATIVACAO_ITENS", 2, raising=False)
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), est(viajante("S1", C, B)), [A, B, C])
    assert erro.value.code == "LISTA_GRANDE"


def test_base_de_estruturas_sem_coluna_cpf():
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A), est(viajante("S1", C, B)).drop(columns=["CPF"]), [A])
    assert erro.value.code == "BASE_SEM_COLUNA"
    assert "CPF" in erro.value.message


def test_base_de_cadastro_sem_coluna_cpf():
    with pytest.raises(InativacaoError) as erro:
        _analisar(cad(USR_A).drop(columns=["CPF"]), est(viajante("S1", C, B)), ["Ana Souza"])
    assert erro.value.code == "BASE_SEM_COLUNA"


def test_usuario_fora_da_lista_nao_e_tocado():
    a = _analisar(cad(USR_A, USR_C), est(viajante("S1", D, A, C)), [A])
    assert a.cpfs == {A}
    assert a.orfas == frozenset()
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `python -m pytest backend/tests/test_inactivation_cascade_service.py backend/tests/test_inativacao_search.py -q`
Expected: FAIL — `ModuleNotFoundError: backend.services.inactivation_cascade_service` e o teste do zero à esquerda com `found == False`.

- [ ] **Step 4: Limites em `Settings` e zero à esquerda em `search_matches`**

Em `backend/core/config.py`, logo abaixo de `MAX_CONTENT_LENGTH: int = 16 * 1024 * 1024`:

```python
    # Inativação em cascata: teto de upload próprio (duas planilhas na mesma requisição) e de itens na lista.
    INATIVACAO_MAX_CONTENT_LENGTH: int = 32 * 1024 * 1024
    MAX_INATIVACAO_ITENS: int = 500
```

Em `backend/services/inactivation_service.py`:

1. Importe `clean_cpf`: acrescente `from backend.shared.cpf_utils import clean_cpf` (ordem alfabética entre os imports `backend.*`).
2. Em `search_matches`, troque
   `frame["CPFdigits"] = frame[cpf_col].apply(lambda v: re.sub(r"\D", "", str(v))) if cpf_col else ""`
   por
   `frame["CPFdigits"] = frame[cpf_col].apply(clean_cpf) if cpf_col else ""`
   (`clean_cpf` restaura o zero à esquerda que o Excel perde, como o motor da ficha já faz).

- [ ] **Step 5: Implement `InactivationCascadeService.analisar`**

`backend/services/inactivation_cascade_service.py`:

```python
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
    "Não foi possível mapear a Estrutura de Aprovação: Usuário encontrado no cadastro, "
    "mas não possui CPF registrado."
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


def _resolver_usuarios(
    busca: dict[str, Any], itens: list[str], escolhidos: set[str]
) -> tuple[list[dict[str, Any]], set[str]]:
    """Transforma o resultado de `search_matches` em usuários com situação; devolve também os CPFs executáveis."""
    digitados_cpf, digitados_email, digitados_nome = _classificar(itens)
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
        if len(grupo) == 1:
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
        usuarios, cpfs = _resolver_usuarios(busca, lista, escolhidos)

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
        digital = impressao_digital(
            {"cpfs": cpfs, "excluidas": excluidas, "compactadas": compactadas, "orfas": orfas}
        )
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
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `python -m pytest backend/tests/test_inactivation_cascade_service.py backend/tests/test_inativacao_search.py -q`
Expected: PASS. Se `test_cpf_repetido_na_lista_entra_uma_vez` falhar por `duplicados`, confira que `search_matches` devolve `duplicates` com o CPF repetido (`[A]`); o serviço só repassa.

Run: `python -m ruff check backend && python -m mypy backend`
Expected: sem erros (ajuste tipos/imports nos arquivos novos se houver).

- [ ] **Step 7: Commit**

```bash
git add backend/core/config.py backend/services/inactivation_service.py backend/services/inactivation_cascade_service.py backend/tests/_cascade_fixtures.py backend/tests/test_inactivation_cascade_service.py backend/tests/test_inativacao_search.py
git commit -m "feat(backend): análise de impacto da inativação em cascata"
```

### Task 4: Execução em cascata e ZIP (`executar` + `ExportService.to_zip_bytes`)

**Files:**
- Modify: `backend/services/inactivation_cascade_service.py` (`Execucao`, `executar`), `backend/services/export_service.py` (`to_zip_bytes`)
- Test: `backend/tests/test_inactivation_executar.py`

**Interfaces:**
- Consumes: Task 3 (`analisar`, `Analise`, `InativacaoError`), Task 2 (`remove_cpfs_and_compact`, `delete_structures`), `InactivationService.process_from_dataframes(df_base, df_lista) -> (DataFrame, dict)`.
- Produces:
  - `Execucao` (dataclass imutável): `ficha: DataFrame` (`Operacao=DELETE`), `estruturas: DataFrame` (`Operacao` na 1ª coluna; `DELETE` nas excluídas, `UPDATE` nas linhas alteradas, vazio nas demais linhas das estruturas compactadas), `resumo: dict` (`usuariosInativados`, `estruturasExcluidas`, `estruturasCompactadas`, `estruturasOrfas`, `linhasEstruturas`).
  - `InactivationCascadeService.executar(df_cadastro, df_estruturas, cpfs: list[str], impressao_recebida: str, ignore_orphan_warning: bool = False) -> Execucao` — levanta `InativacaoError` com `NADA_A_EXECUTAR` (400), `ANALISE_DIVERGENTE` (409) ou `ORFAS_SEM_CONFIRMACAO` (400, `extra={"estruturasOrfas": [...]}`), nessa ordem de checagem.
  - `ExportService.to_zip_bytes(files: dict[str, io.BytesIO]) -> io.BytesIO`.

- [ ] **Step 1: Write the failing tests**

`backend/tests/test_inactivation_executar.py`:

```python
import io
import zipfile

import pytest
from _cascade_fixtures import A, B, C, D, USR_A, cad, est, viajante

from backend.services.export_service import ExportService
from backend.services.inactivation_cascade_service import InactivationCascadeService, InativacaoError


def _digital(cadastro, estruturas, cpfs):
    return InactivationCascadeService.analisar(cadastro, estruturas, cpfs).payload["impressaoDigital"]


def _executar(cadastro, estruturas, cpfs, **kw):
    return InactivationCascadeService.executar(cadastro, estruturas, cpfs, _digital(cadastro, estruturas, cpfs), **kw)


def test_gera_ficha_delete_e_estruturas_com_delete_e_update():
    estruturas = est(viajante("S1", A, B), viajante("S2", C, B, A))
    r = _executar(cad(USR_A), estruturas, [A])
    assert r.ficha["Operacao"].tolist() == ["DELETE"]
    assert dict(zip(r.estruturas["AprovacaoId"], r.estruturas["Operacao"])) == {"S1": "DELETE", "S2": "UPDATE"}
    s2 = r.estruturas[r.estruturas["AprovacaoId"] == "S2"].iloc[0]
    assert (s2["LoginAprovador_1"], s2["LoginAprovador_2"]) == (B, "")
    assert r.estruturas.columns[0] == "Operacao"
    assert r.resumo == {
        "usuariosInativados": 1,
        "estruturasExcluidas": 1,
        "estruturasCompactadas": 1,
        "estruturasOrfas": 0,
        "linhasEstruturas": 2,
    }


def test_estrutura_orfa_exige_confirmacao():
    cadastro, estruturas = cad(USR_A), est(viajante("S3", C, A))
    with pytest.raises(InativacaoError) as erro:
        _executar(cadastro, estruturas, [A])
    assert erro.value.code == "ORFAS_SEM_CONFIRMACAO"
    assert erro.value.status == 400
    assert erro.value.extra["estruturasOrfas"] == ["S3"]

    r = _executar(cadastro, estruturas, [A], ignore_orphan_warning=True)
    s3 = r.estruturas.iloc[0]
    assert (s3["AprovacaoId"], s3["Operacao"], s3["LoginAprovador_1"]) == ("S3", "UPDATE", "")


def test_impressao_digital_divergente_e_recusada():
    with pytest.raises(InativacaoError) as erro:
        InactivationCascadeService.executar(cad(USR_A), est(viajante("S1", A, B)), [A], "0" * 64)
    assert erro.value.code == "ANALISE_DIVERGENTE"
    assert erro.value.status == 409


def test_nenhum_executavel_e_recusado_antes_de_comparar_a_impressao():
    with pytest.raises(InativacaoError) as erro:
        InactivationCascadeService.executar(cad(USR_A), est(viajante("S1", C, B)), [D], "qualquer")
    assert erro.value.code == "NADA_A_EXECUTAR"
    with pytest.raises(InativacaoError) as erro:
        InactivationCascadeService.executar(cad(USR_A), est(viajante("S1", C, B)), [], "qualquer")
    assert erro.value.code == "NADA_A_EXECUTAR"


def test_sem_estruturas_afetadas_o_arquivo_de_estruturas_sai_so_com_cabecalho():
    r = _executar(cad(USR_A), est(viajante("S1", C, B)), [A])
    assert len(r.ficha) == 1
    assert r.estruturas.empty
    assert list(r.estruturas.columns[:2]) == ["Operacao", "AprovacaoId"]


def test_execucao_repetida_da_a_mesma_saida():
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))
    primeira, segunda = _executar(cadastro, estruturas, [A]), _executar(cadastro, estruturas, [A])
    assert primeira.ficha.equals(segunda.ficha)
    assert primeira.estruturas.equals(segunda.estruturas)


def test_nao_altera_as_entradas():
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))
    antes_c, antes_e = cadastro.copy(), estruturas.copy()
    _executar(cadastro, estruturas, [A])
    assert cadastro.equals(antes_c)
    assert estruturas.equals(antes_e)


def test_to_zip_bytes_guarda_cada_arquivo_com_o_nome_dado():
    saida = ExportService.to_zip_bytes({"a.xlsx": io.BytesIO(b"1"), "b.xlsx": io.BytesIO(b"2")})
    with zipfile.ZipFile(saida) as z:
        assert z.namelist() == ["a.xlsx", "b.xlsx"]
        assert z.read("b.xlsx") == b"2"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest backend/tests/test_inactivation_executar.py -q`
Expected: FAIL — `AttributeError: ... has no attribute 'executar'` / `to_zip_bytes`.

- [ ] **Step 3: Implement `to_zip_bytes`**

Em `backend/services/export_service.py`, acrescente `import zipfile` logo abaixo de `import io` e, dentro de `class ExportService`, ao final:

```python
    @staticmethod
    def to_zip_bytes(files: dict[str, io.BytesIO]) -> io.BytesIO:
        """Empacota arquivos já gerados (nome -> bytes) num ZIP em memória."""
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in files.items():
                zf.writestr(name, data.getvalue())
        output.seek(0)
        return output
```

- [ ] **Step 4: Implement `executar`**

Em `backend/services/inactivation_cascade_service.py`, logo abaixo da classe `Analise` acrescente:

```python
@dataclass(frozen=True)
class Execucao:
    ficha: pd.DataFrame  # ficha de inativação (Operacao=DELETE)
    estruturas: pd.DataFrame  # estruturas afetadas: DELETE nas excluídas, UPDATE nas alteradas
    resumo: dict[str, Any]
```

e, dentro de `InactivationCascadeService`, depois de `analisar`:

```python
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
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest backend/tests/test_inactivation_executar.py backend/tests/test_inactivation_cascade_service.py -q`
Expected: PASS. Se `test_gera_ficha_delete...` falhar em `ficha["Operacao"]`, confira `MODEL_COLS` em `backend/domain/` (a coluna `Operacao` é escrita por `processar_inativacao_from_paths`).

Run: `python -m ruff check backend && python -m mypy backend`
Expected: sem erros.

- [ ] **Step 6: Commit**

```bash
git add backend/services/inactivation_cascade_service.py backend/services/export_service.py backend/tests/test_inactivation_executar.py
git commit -m "feat(backend): execução da inativação em cascata e empacotamento em ZIP"
```

### Task 5: Rotas `/analisar` e `/executar`, auditoria e remoção das rotas antigas

**Files:**
- Rewrite: `backend/api/inativacao.py`
- Create: `backend/tests/test_inativacao_analisar_api.py`, `backend/tests/test_inativacao_executar_api.py`
- Delete: `backend/tests/test_inativacao_api.py`, `backend/tests/test_inativacao_buscar_api.py`
- Modify (só os testes de rota removida): `backend/tests/test_inativacao_generation_api.py`, `backend/tests/test_dead_params_removed.py`, `backend/tests/test_upload_dir.py`

**Interfaces:**
- Consumes: Tasks 1, 3 e 4 (`InactivationCascadeService.analisar/executar`, `InativacaoError`, `ExportService.to_zip_bytes`, `mascarar_cpf`).
- Produces (HTTP, multipart):
  - `POST /api/inativacao/analisar` — campos `cadastro`, `estruturas` (arquivos), lista por `itens` (JSON), `lista` (arquivo) ou `lista_text`, e opcional `selecionados` (JSON de CPFs). `200` = `Analise.payload`.
  - `POST /api/inativacao/executar` — `cadastro`, `estruturas`, `cpfs` (JSON), `impressaoDigital`, `ignore_orphan_warning` (`true|false`). `200` = ZIP `inativacao.zip` com `saida_inativacao.xlsx` e `estruturas_atualizadas.xlsx`.
  - Erros: `{"error": str, "code": str, ...extra}` com o status de `InativacaoError`; `413 ARQUIVO_GRANDE`; `500 ERRO_INTERNO`.
  - Eventos de auditoria `inativacao_analise` e `inativacao_execucao` (só contagens, impressão digital e CPF mascarado).

- [ ] **Step 1: Write the failing tests — `/analisar`**

`backend/tests/test_inativacao_analisar_api.py`:

```python
import io
import json
import os

import pandas as pd
from _cascade_fixtures import A, B, C, USR_A, cad, est, viajante
from _helpers import xlsx_upload

from backend.core.config import settings
from backend.services.audit_service import AuditService
from backend.services.inactivation_cascade_service import InactivationCascadeService

ROTA = "/api/inativacao/analisar"


def _post(client, cadastro=None, estruturas=None, **campos):
    data = dict(campos)
    if cadastro is not None:
        data["cadastro"] = xlsx_upload(cadastro, "cadastro.xlsx")
    if estruturas is not None:
        data["estruturas"] = xlsx_upload(estruturas, "estruturas.xlsx")
    return client.post(ROTA, data=data, content_type="multipart/form-data")


def _bases():
    return cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))


def test_analisa_por_itens_json(client):
    resp = _post(client, *_bases(), itens=json.dumps([A]))
    assert resp.status_code == 200, resp.get_data(as_text=True)
    corpo = resp.get_json()
    assert set(corpo) == {"usuarios", "resumo", "impressaoDigital"}
    assert corpo["usuarios"][0]["situacao"] == "EXECUTAVEL"
    assert corpo["usuarios"][0]["estruturasViajante"] == ["S1"]
    assert corpo["resumo"]["estruturasCompactadas"] == 1


def test_analisa_por_lista_text(client):
    resp = _post(client, *_bases(), lista_text=f"{A}\n")
    assert resp.status_code == 200
    assert resp.get_json()["resumo"]["executaveis"] == 1


def test_analisa_por_lista_em_arquivo(client):
    lista = xlsx_upload(pd.DataFrame([{"CPF": A}]), "lista.xlsx")
    resp = _post(client, *_bases(), lista=lista)
    assert resp.status_code == 200
    assert resp.get_json()["resumo"]["executaveis"] == 1


def test_selecionados_resolvem_homonimos(client):
    cadastro = cad((A, "Joao Silva", "j1@x.com", "ATIVO"), (B, "Joao Silva", "j2@x.com", "ATIVO"))
    estruturas = est(viajante("S9", C, A))
    pendente = _post(client, cadastro, estruturas, itens=json.dumps(["Joao Silva"])).get_json()
    assert pendente["usuarios"][0]["situacao"] == "PENDENTE_SELECAO"
    escolhido = _post(client, cadastro, estruturas, itens=json.dumps(["Joao Silva"]), selecionados=json.dumps([B]))
    assert escolhido.get_json()["usuarios"][0]["situacao"] == "EXECUTAVEL"


def test_sem_a_base_de_estruturas(client):
    resp = _post(client, cad(USR_A), None, itens=json.dumps([A]))
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "BASE_AUSENTE"


def test_arquivo_que_nao_e_excel(client):
    data = {
        "cadastro": (io.BytesIO(b"isto nao e um excel"), "cadastro.xlsx"),
        "estruturas": xlsx_upload(_bases()[1], "estruturas.xlsx"),
        "itens": json.dumps([A]),
    }
    resp = client.post(ROTA, data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "ARQUIVO_INVALIDO"


def test_lista_vazia(client):
    resp = _post(client, *_bases())
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "LISTA_VAZIA"


def test_base_de_estruturas_sem_coluna_cpf(client):
    resp = _post(client, cad(USR_A), _bases()[1].drop(columns=["CPF"]), itens=json.dumps([A]))
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "BASE_SEM_COLUNA"


def test_arquivo_acima_do_teto_da_rota(client, monkeypatch):
    monkeypatch.setattr(settings, "INATIVACAO_MAX_CONTENT_LENGTH", 200, raising=False)
    resp = _post(client, *_bases(), itens=json.dumps([A]))
    assert resp.status_code == 413
    assert resp.get_json()["code"] == "ARQUIVO_GRANDE"


def test_apaga_os_temporarios_no_sucesso_e_no_erro(client):
    _post(client, *_bases(), itens=json.dumps([A]))
    _post(client, *_bases())
    assert os.listdir(settings.UPLOAD_FOLDER) == []


def test_auditoria_nao_guarda_dado_pessoal(client):
    assert _post(client, *_bases(), itens=json.dumps([A])).status_code == 200
    eventos = AuditService.list_events()
    assert any(e["event_type"] == "inativacao_analise" and e["status"] == "success" for e in eventos)
    texto = json.dumps(eventos, ensure_ascii=False)
    for pessoal in (A, "Ana Souza", "ana@x.com"):
        assert pessoal not in texto
    assert "***." in texto


def test_erro_interno_vira_500_sem_vazar_detalhe(client, monkeypatch):
    def boom(*args, **kwargs):
        raise RuntimeError("segredo interno")

    monkeypatch.setattr(InactivationCascadeService, "analisar", boom)
    resp = _post(client, *_bases(), itens=json.dumps([A]))
    assert resp.status_code == 500
    assert resp.get_json()["code"] == "ERRO_INTERNO"
    assert "segredo" not in resp.get_data(as_text=True)
```

- [ ] **Step 2: Write the failing tests — `/executar`**

`backend/tests/test_inativacao_executar_api.py`:

```python
import io
import json
import os
import zipfile

import pandas as pd
from _cascade_fixtures import A, B, C, USR_A, cad, est, viajante
from _helpers import xlsx_upload

from backend.core.config import settings
from backend.services.audit_service import AuditService
from backend.services.inactivation_cascade_service import InactivationCascadeService

ROTA = "/api/inativacao/executar"


def _digital(cadastro, estruturas, cpfs):
    return InactivationCascadeService.analisar(cadastro, estruturas, cpfs).payload["impressaoDigital"]


def _post(client, cadastro, estruturas, cpfs, digital=None, **campos):
    data = {
        "cadastro": xlsx_upload(cadastro, "cadastro.xlsx"),
        "estruturas": xlsx_upload(estruturas, "estruturas.xlsx"),
        "cpfs": json.dumps(cpfs),
        "impressaoDigital": digital if digital is not None else _digital(cadastro, estruturas, cpfs),
    }
    data.update(campos)
    return client.post(ROTA, data=data, content_type="multipart/form-data")


def _planilha(zf, nome):
    return pd.read_excel(io.BytesIO(zf.read(nome)), dtype=str).fillna("")


def test_devolve_zip_com_as_duas_planilhas(client):
    cadastro, estruturas = cad(USR_A), est(viajante("S1", A, B), viajante("S2", C, B, A))
    resp = _post(client, cadastro, estruturas, [A])
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert resp.headers["Content-Type"].startswith("application/zip")
    with zipfile.ZipFile(io.BytesIO(resp.data)) as zf:
        assert zf.namelist() == ["saida_inativacao.xlsx", "estruturas_atualizadas.xlsx"]
        assert _planilha(zf, "saida_inativacao.xlsx")["Operacao"].tolist() == ["DELETE"]
        est_saida = _planilha(zf, "estruturas_atualizadas.xlsx")
        assert dict(zip(est_saida["AprovacaoId"], est_saida["Operacao"])) == {"S1": "DELETE", "S2": "UPDATE"}


def test_grava_auditoria_sem_dado_pessoal(client):
    resp = _post(client, cad(USR_A), est(viajante("S1", A, B)), [A])
    assert resp.status_code == 200
    eventos = AuditService.list_events()
    assert any(e["event_type"] == "inativacao_execucao" and e["status"] == "success" for e in eventos)
    texto = json.dumps(eventos, ensure_ascii=False)
    for pessoal in (A, "Ana Souza", "ana@x.com"):
        assert pessoal not in texto


def test_impressao_digital_divergente_da_409(client):
    resp = _post(client, cad(USR_A), est(viajante("S1", A, B)), [A], digital="0" * 64)
    assert resp.status_code == 409
    assert resp.get_json()["code"] == "ANALISE_DIVERGENTE"


def test_estrutura_orfa_pede_confirmacao_e_depois_executa(client):
    cadastro, estruturas = cad(USR_A), est(viajante("S3", C, A))
    resp = _post(client, cadastro, estruturas, [A])
    assert resp.status_code == 400
    corpo = resp.get_json()
    assert corpo["code"] == "ORFAS_SEM_CONFIRMACAO"
    assert corpo["estruturasOrfas"] == ["S3"]
    assert _post(client, cadastro, estruturas, [A], ignore_orphan_warning="true").status_code == 200


def test_nada_a_executar(client):
    resp = _post(client, cad(USR_A), est(viajante("S1", C, B)), [], digital="x")
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "NADA_A_EXECUTAR"


def test_sem_arquivos_nunca_responde_sucesso(client):
    resp = client.post(ROTA, data={"cpfs": json.dumps([A])}, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.get_json()["code"] == "BASE_AUSENTE"


def test_arquivo_acima_do_teto_da_rota(client, monkeypatch):
    monkeypatch.setattr(settings, "INATIVACAO_MAX_CONTENT_LENGTH", 200, raising=False)
    resp = _post(client, cad(USR_A), est(viajante("S1", A, B)), [A])
    assert resp.status_code == 413
    assert resp.get_json()["code"] == "ARQUIVO_GRANDE"


def test_apaga_os_temporarios(client):
    _post(client, cad(USR_A), est(viajante("S1", A, B)), [A])
    _post(client, cad(USR_A), est(viajante("S1", A, B)), [A], digital="0" * 64)
    assert os.listdir(settings.UPLOAD_FOLDER) == []


def test_erro_interno_vira_500_e_nunca_sucesso(client, monkeypatch):
    def boom(*args, **kwargs):
        raise RuntimeError("bug interno")

    monkeypatch.setattr(InactivationCascadeService, "executar", boom)
    resp = _post(client, cad(USR_A), est(viajante("S1", A, B)), [A], digital="x")
    assert resp.status_code == 500
    assert resp.get_json()["code"] == "ERRO_INTERNO"


def test_rotas_antigas_foram_removidas(client):
    for rota in ("/api/process_inativacao", "/api/preview_inativacao", "/api/inativacao/buscar"):
        assert client.post(rota, data={}, content_type="multipart/form-data").status_code == 404
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `python -m pytest backend/tests/test_inativacao_analisar_api.py backend/tests/test_inativacao_executar_api.py -q`
Expected: FAIL — as rotas novas devolvem 404 e as antigas ainda existem.

- [ ] **Step 4: Rewrite `backend/api/inativacao.py`**

Substitua o arquivo inteiro por (as rotas `/inativacao/buscar`, `/process_inativacao` e `/preview_inativacao` saem; a busca continua no `InactivationService`):

```python
import json
import os

import pandas as pd
from flask import Blueprint, jsonify, request, send_file
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService
from backend.services.export_service import ExportService
from backend.services.inactivation_cascade_service import Analise, InactivationCascadeService, InativacaoError
from backend.services.inactivation_service import InactivationService
from backend.shared.cpf_mask import mascarar_cpf
from backend.shared.upload_validation import save_and_validate_upload

logger = get_logger()

inativacao_bp = Blueprint("inativacao", __name__, url_prefix="/api")

_VERDADEIRO = {"1", "true", "yes", "on"}


@inativacao_bp.before_request
def _limite_de_upload() -> None:
    # Duas planilhas grandes na mesma requisição estouram os 16 MB globais: teto próprio destas rotas.
    request.max_content_length = settings.INATIVACAO_MAX_CONTENT_LENGTH


@inativacao_bp.errorhandler(RequestEntityTooLarge)
def _arquivo_grande(_exc):
    limite = max(1, settings.INATIVACAO_MAX_CONTENT_LENGTH // (1024 * 1024))
    mensagem = f"Os arquivos enviados passam do limite de {limite} MB."
    return jsonify({"error": mensagem, "code": "ARQUIVO_GRANDE"}), 413


def _erro(exc: InativacaoError):
    return jsonify({"error": exc.message, "code": exc.code, **exc.extra}), exc.status


def _limpar(paths: list[str]) -> None:
    for path in paths:
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception as exc:  # pragma: no cover - limpeza nunca derruba a resposta
            logger.warning("Falha ao remover temporário %s: %s", path, exc)


def _ler_planilha(campo: str, rotulo: str, paths: list[str]) -> pd.DataFrame:
    arquivo = request.files.get(campo)
    if not arquivo:
        raise InativacaoError("BASE_AUSENTE", f"Envie a {rotulo}.")
    path, err = save_and_validate_upload(arquivo, settings.UPLOAD_FOLDER, label=rotulo)
    if path:
        paths.append(path)
    if err:
        raise InativacaoError("ARQUIVO_INVALIDO", err)
    try:
        return pd.read_excel(path, dtype=str).fillna("")
    except Exception:
        raise InativacaoError("ARQUIVO_INVALIDO", f"{rotulo}: não foi possível ler a planilha.") from None


def _json_lista(campo: str) -> list[str]:
    bruto = request.form.get(campo, "")
    if not bruto:
        return []
    try:
        dados = json.loads(bruto)
    except ValueError:
        raise InativacaoError("LISTA_VAZIA", f"O campo '{campo}' não é um JSON válido.") from None
    return [str(x).strip() for x in dados if str(x).strip()] if isinstance(dados, list) else []


def _extrair_itens(paths: list[str]) -> list[str]:
    """Itens da lista: `itens` (JSON), `lista` (planilha) ou `lista_text` (um por linha), nessa ordem."""
    if request.form.get("itens"):
        return _json_lista("itens")
    arquivo = request.files.get("lista")
    if arquivo:
        path, err = save_and_validate_upload(arquivo, settings.UPLOAD_FOLDER, label="lista")
        if path:
            paths.append(path)
        if err:
            raise InativacaoError("ARQUIVO_INVALIDO", err)
        df = InactivationService.normalize_lista_columns(pd.read_excel(path, dtype=str).fillna(""))
        itens: list[str] = []
        for _idx, row in df.iterrows():
            valores = (str(row.get(col, "")).strip() for col in ("CPF", "Email", "NomeCompleto"))
            itens.append(next((v for v in valores if v), ""))
        return [i for i in itens if i]
    return [linha.strip() for linha in request.form.get("lista_text", "").split("\n") if linha.strip()]


def _detalhes_analise(analise: Analise) -> dict:
    """Só contagens, impressão digital e CPF mascarado: nome e e-mail nunca vão para o histórico."""
    return {
        "resumo": analise.payload["resumo"],
        "impressaoDigital": analise.payload["impressaoDigital"],
        "usuarios": [{"cpf": u["cpfMascarado"], "situacao": u["situacao"]} for u in analise.payload["usuarios"]],
    }


def _interno(evento: str, rota: str):
    logger.exception("Erro em %s", rota)
    AuditService.record(event_type=evento, status="error", details={"code": "ERRO_INTERNO"})
    return jsonify({"error": "Erro interno ao processar a solicitação.", "code": "ERRO_INTERNO"}), 500


@inativacao_bp.route("/inativacao/analisar", methods=["POST"])
def api_inativacao_analisar():
    paths: list[str] = []
    try:
        df_cadastro = _ler_planilha("cadastro", "base de cadastro", paths)
        df_estruturas = _ler_planilha("estruturas", "base de estruturas", paths)
        itens = _extrair_itens(paths)
        analise = InactivationCascadeService.analisar(df_cadastro, df_estruturas, itens, _json_lista("selecionados"))
        AuditService.record(event_type="inativacao_analise", status="success", details=_detalhes_analise(analise))
        return jsonify(analise.payload), 200
    except InativacaoError as exc:
        AuditService.record(event_type="inativacao_analise", status="error", details={"code": exc.code})
        return _erro(exc)
    except HTTPException:
        raise
    except Exception:
        return _interno("inativacao_analise", "/api/inativacao/analisar")
    finally:
        _limpar(paths)


@inativacao_bp.route("/inativacao/executar", methods=["POST"])
def api_inativacao_executar():
    paths: list[str] = []
    try:
        df_cadastro = _ler_planilha("cadastro", "base de cadastro", paths)
        df_estruturas = _ler_planilha("estruturas", "base de estruturas", paths)
        cpfs = _json_lista("cpfs")
        ignorar_orfas = request.form.get("ignore_orphan_warning", "").lower() in _VERDADEIRO
        execucao = InactivationCascadeService.executar(
            df_cadastro, df_estruturas, cpfs, request.form.get("impressaoDigital", ""), ignorar_orfas
        )
        arquivos = {
            "saida_inativacao.xlsx": ExportService.to_excel_bytes(execucao.ficha, sheet_name="Inativacao"),
            "estruturas_atualizadas.xlsx": ExportService.to_excel_bytes(execucao.estruturas, sheet_name="Aprovacao"),
        }
        pacote = ExportService.to_zip_bytes(arquivos)
        AuditService.record(
            event_type="inativacao_execucao",
            status="success",
            details={"resumo": execucao.resumo, "cpfs": [mascarar_cpf(c) for c in cpfs]},
        )
        return send_file(pacote, download_name="inativacao.zip", as_attachment=True, mimetype="application/zip")
    except InativacaoError as exc:
        AuditService.record(event_type="inativacao_execucao", status="error", details={"code": exc.code})
        return _erro(exc)
    except HTTPException:
        raise
    except Exception:
        return _interno("inativacao_execucao", "/api/inativacao/executar")
    finally:
        _limpar(paths)
```

- [ ] **Step 5: Remove os testes das rotas que saíram**

```bash
git rm backend/tests/test_inativacao_api.py backend/tests/test_inativacao_buscar_api.py
```

(`test_inativacao_api.py` inclui `test_inativacao_executar_gone`, que garantia a ausência da rota falsa; a proteção contra o sucesso falso agora está em `test_sem_arquivos_nunca_responde_sucesso` e `test_erro_interno_vira_500_e_nunca_sucesso`. A cobertura da leitura da lista por JSON, texto e arquivo passou para `test_inativacao_analisar_api.py`.)

Edições cirúrgicas (os testes do motor `processar_inativacao_from_paths` **ficam**):

- `backend/tests/test_inativacao_generation_api.py`: apague do `def test_process_inativacao_endpoint_returns_xlsx` até o fim do arquivo (3 testes de rota) e troque `from _helpers import valid_cpf, xlsx_upload` por `from _helpers import valid_cpf`. Restam os 4 testes do motor.
- `backend/tests/test_dead_params_removed.py`: apague `test_process_inativacao_still_works` e os imports que ficarem sem uso (`pandas`, `valid_cpf`, `xlsx_upload`); `test_signatures_have_no_fuzzy_params` fica.
- `backend/tests/test_upload_dir.py`: apague `test_process_inativacao_cleans_temp_files` (a limpeza agora é testada em `test_apaga_os_temporarios*`); mantenha os outros dois testes.

- [ ] **Step 6: Run the whole backend suite and linters**

Run: `python -m pytest backend/tests -q`
Expected: PASS (nenhum teste do fluxo antigo restante).

Run: `python -m ruff check backend && python -m ruff format --check backend && python -m mypy backend`
Expected: sem erros. Se o `format --check` reclamar dos arquivos novos, rode `python -m ruff format` neles; se o ruff apontar imports sem uso nos testes editados, remova só esses.

- [ ] **Step 7: Commit**

```bash
git add -A backend
git commit -m "feat(backend): rotas /inativacao/analisar e /executar com ZIP e auditoria sem dado pessoal; remove as rotas antigas"
```

---

### Task 6: Propriedades e volume

**Files:**
- Test: `backend/tests/test_inactivation_volume.py`

**Interfaces:**
- Consumes: `InactivationCascadeService.analisar/executar` (Tasks 3 e 4).

- [ ] **Step 1: Write the test**

`backend/tests/test_inactivation_volume.py`:

```python
import time

import pandas as pd
from _helpers import valid_cpf

from backend.services.inactivation_cascade_service import InactivationCascadeService

LINHAS = 50_000
USUARIOS = 100
TETO_SEGUNDOS = 15


def _bases():
    inativar = [valid_cpf(i) for i in range(1, USUARIOS + 1)]
    outros = [valid_cpf(i) for i in range(1000, 1030)]
    linhas = [
        {
            "AprovacaoId": f"S{i}",
            "AprovacaoPor": "VIAJANTE" if i % 10 == 0 else "CCEMPRESA",
            "CPF": inativar[(i // 10) % USUARIOS] if i % 10 == 0 else "",
            "LoginAprovador_1": outros[i % 30],
            "LoginAprovador_2": inativar[i % USUARIOS] if i % 3 == 0 else "",
            "LoginAprovador_3": outros[(i + 1) % 30],
        }
        for i in range(LINHAS)
    ]
    cadastro = pd.DataFrame(
        [
            {"CPF": cpf, "NomeCompleto": f"Usuario {n} Teste", "Email": f"u{n}@x.com", "Status": "ATIVO"}
            for n, cpf in enumerate(inativar)
        ]
    )
    return cadastro, pd.DataFrame(linhas).astype(str), inativar


def test_analise_e_execucao_de_50_mil_linhas_com_100_usuarios_cabem_no_teto():
    cadastro, estruturas, cpfs = _bases()

    inicio = time.perf_counter()
    analise = InactivationCascadeService.analisar(cadastro, estruturas, cpfs)
    duracao_analise = time.perf_counter() - inicio
    assert analise.payload["resumo"]["executaveis"] == USUARIOS
    assert duracao_analise < TETO_SEGUNDOS

    inicio = time.perf_counter()
    execucao = InactivationCascadeService.executar(
        cadastro, estruturas, cpfs, analise.payload["impressaoDigital"], ignore_orphan_warning=True
    )
    duracao_execucao = time.perf_counter() - inicio
    assert execucao.resumo["usuariosInativados"] == USUARIOS
    assert duracao_execucao < TETO_SEGUNDOS
```

- [ ] **Step 2: Run it and measure**

Run: `python -m pytest backend/tests/test_inactivation_volume.py -q --durations=3`
Expected: PASS. Anote os tempos: se algum passar de ~5 s, o gargalo provável é o laço de `remove_cpfs_and_compact` (`df_out.at` por célula) — reduza atribuindo colunas inteiras por linha (`df_out.loc[idx, approver_cols] = valores`) **mantendo** a suíte `test_approval_multi_cpf.py` verde; se ficarem muito abaixo do teto, aperte `TETO_SEGUNDOS` para ~3× o medido.

- [ ] **Step 3: Commit**

```bash
git add backend/tests/test_inactivation_volume.py
git commit -m "test(backend): volume de 50 mil linhas com 100 usuários na inativação em cascata"
```

---

### Task 7: Camada de API do frontend (`ApiRejection.code`, `inativacaoApi.ts`)

**Files:**
- Modify: `frontend-react/src/lib/api.ts`
- Create: `frontend-react/src/lib/inativacaoApi.ts`
- Test: `frontend-react/src/lib/inativacaoApi.test.ts`

**Interfaces:**
- Produces (`lib/inativacaoApi.ts`):
  - tipos `Situacao`, `Candidato`, `EstruturaComoAprovador`, `UsuarioAnalise`, `ResumoAnalise`, `AnaliseInativacao` (espelham o JSON de `/analisar`);
  - `class InativacaoApiError extends Error { readonly code: string }`;
  - `postAnalisar(cadastro: File, estruturas: File, itens: string[], selecionados: string[]): Promise<AnaliseInativacao>`;
  - `postExecutar(cadastro: File, estruturas: File, cpfs: string[], impressaoDigital: string, ignorarOrfas: boolean, onProgress?: (pct: number) => void): Promise<Blob>`.
- `ApiRejection` ganha `readonly code?: string` (3º argumento do construtor).

- [ ] **Step 1: Write the failing tests**

`frontend-react/src/lib/inativacaoApi.test.ts`:

```ts
import { afterEach, describe, expect, it, vi } from "vitest";
import { InativacaoApiError, postAnalisar, postExecutar, type AnaliseInativacao } from "./inativacaoApi";

const arquivo = (nome: string) => new File(["x"], nome);

const ANALISE: AnaliseInativacao = {
  usuarios: [],
  resumo: { executaveis: 0, estruturasExcluidas: 0, estruturasCompactadas: 0, estruturasOrfas: 0, duplicados: [] },
  impressaoDigital: "abc",
};

class FakeXhr {
  static respostas: { status: number; corpo: unknown }[] = [];
  static enviados: FormData[] = [];
  upload = { addEventListener: () => {} };
  responseType = "";
  status = 0;
  response: Blob | null = null;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  open() {}
  send(formData: FormData) {
    FakeXhr.enviados.push(formData);
    const proxima = FakeXhr.respostas.shift();
    queueMicrotask(() => {
      this.status = proxima?.status ?? 500;
      const corpo = proxima?.corpo ?? "";
      this.response = new Blob([typeof corpo === "string" ? corpo : JSON.stringify(corpo)]);
      this.onload?.();
    });
  }
}

afterEach(() => {
  vi.unstubAllGlobals();
  FakeXhr.respostas = [];
  FakeXhr.enviados = [];
});

describe("postAnalisar", () => {
  it("envia as bases, a lista e as escolhas, e devolve a análise", async () => {
    const fetchMock = vi.fn().mockResolvedValue({ ok: true, status: 200, json: async () => ANALISE });
    vi.stubGlobal("fetch", fetchMock);

    const out = await postAnalisar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["123"], ["456"]);

    expect(out).toEqual(ANALISE);
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe("/api/inativacao/analisar");
    const fd = init.body as FormData;
    expect(fd.get("itens")).toBe(JSON.stringify(["123"]));
    expect(fd.get("selecionados")).toBe(JSON.stringify(["456"]));
    expect((fd.get("cadastro") as File).name).toBe("c.xlsx");
    expect((fd.get("estruturas") as File).name).toBe("e.xlsx");
  });

  it("transforma a recusa do servidor em InativacaoApiError com o code", async () => {
    const recusa = { ok: false, status: 400, json: async () => ({ error: "Envie a base de estruturas.", code: "BASE_AUSENTE" }) };
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(recusa));

    const erro = await postAnalisar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["1"], []).catch((e: unknown) => e);

    expect(erro).toBeInstanceOf(InativacaoApiError);
    expect(erro).toMatchObject({ code: "BASE_AUSENTE", message: "Envie a base de estruturas." });
  });

  it("sem corpo JSON usa ERRO_INTERNO e o status na mensagem", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: false, status: 502, json: async () => Promise.reject(new Error("html")) }));
    const erro = await postAnalisar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["1"], []).catch((e: unknown) => e);
    expect(erro).toMatchObject({ code: "ERRO_INTERNO", message: "Falha na requisição (502)" });
  });
});

describe("postExecutar", () => {
  it("envia cpfs, impressão digital e a confirmação das órfãs, e devolve o ZIP", async () => {
    vi.stubGlobal("XMLHttpRequest", FakeXhr);
    FakeXhr.respostas.push({ status: 200, corpo: "conteudo-do-zip" });

    const blob = await postExecutar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["111", "222"], "digital", true);

    expect(blob.size).toBeGreaterThan(0);
    const fd = FakeXhr.enviados[0];
    expect(fd.get("cpfs")).toBe(JSON.stringify(["111", "222"]));
    expect(fd.get("impressaoDigital")).toBe("digital");
    expect(fd.get("ignore_orphan_warning")).toBe("true");
  });

  it("repassa o code do servidor (ex.: análise divergente)", async () => {
    vi.stubGlobal("XMLHttpRequest", FakeXhr);
    FakeXhr.respostas.push({ status: 409, corpo: { error: "A análise mudou.", code: "ANALISE_DIVERGENTE" } });

    const erro = await postExecutar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["1"], "d", false).catch((e: unknown) => e);

    expect(erro).toBeInstanceOf(InativacaoApiError);
    expect(erro).toMatchObject({ code: "ANALISE_DIVERGENTE", message: "A análise mudou." });
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npm --prefix frontend-react test -- src/lib/inativacaoApi.test.ts`
Expected: FAIL — módulo `./inativacaoApi` não existe.

- [ ] **Step 3: `ApiRejection` carrega o `code`**

Em `frontend-react/src/lib/api.ts`, na classe `ApiRejection`:

```ts
export class ApiRejection extends Error {
  /** `{nome do arquivo: motivo}`, quando o servidor separa o problema por arquivo. */
  readonly fileErrors?: Record<string, string>;
  /** Código estável do erro de negócio (`code`), quando o servidor o envia: a tela decide por ele, não pelo texto. */
  readonly code?: string;

  constructor(message: string, fileErrors?: Record<string, string>, code?: string) {
    super(message);
    this.fileErrors = fileErrors;
    this.code = code;
  }
}
```

E em `postFormForBlob`, troque `reject(new ApiRejection(obj.error || `Erro ${xhr.status}`, asFileErrors(obj.errors)));` por:

```ts
          reject(
            new ApiRejection(
              obj.error || `Erro ${xhr.status}`,
              asFileErrors(obj.errors),
              typeof obj.code === "string" ? obj.code : undefined,
            ),
          );
```

- [ ] **Step 4: Create `inativacaoApi.ts`**

`frontend-react/src/lib/inativacaoApi.ts`:

```ts
import { ApiRejection, postFormForBlob } from "./api";

export type Situacao = "EXECUTAVEL" | "SEM_CPF" | "JA_INATIVO" | "NAO_LOCALIZADO" | "PENDENTE_SELECAO";

export interface Candidato {
  cpf: string;
  cpfMascarado: string;
  nome: string;
  email: string;
}

export interface EstruturaComoAprovador {
  aprovacaoId: string;
  posicoes: number[];
  segundoNivel: boolean;
  acao: "COMPACTACAO" | "ORFA";
}

export interface UsuarioAnalise {
  cpf: string | null;
  cpfMascarado: string;
  nome: string;
  email: string;
  situacao: Situacao;
  alerta: string | null;
  estruturasViajante: string[];
  comoAprovador: EstruturaComoAprovador[];
  candidatos: Candidato[];
}

export interface ResumoAnalise {
  executaveis: number;
  estruturasExcluidas: number;
  estruturasCompactadas: number;
  estruturasOrfas: number;
  duplicados: string[];
}

export interface AnaliseInativacao {
  usuarios: UsuarioAnalise[];
  resumo: ResumoAnalise;
  impressaoDigital: string;
}

/** Erro de negócio da inativação: a tela decide pelo `code` estável do servidor, não pelo texto. */
export class InativacaoApiError extends Error {
  readonly code: string;

  constructor(message: string, code: string) {
    super(message);
    this.name = "InativacaoApiError";
    this.code = code;
  }
}

function basesForm(cadastro: File, estruturas: File): FormData {
  const fd = new FormData();
  fd.append("cadastro", cadastro);
  fd.append("estruturas", estruturas);
  return fd;
}

/** Diagnóstico de impacto: não altera nada no servidor. */
export async function postAnalisar(
  cadastro: File,
  estruturas: File,
  itens: string[],
  selecionados: string[],
): Promise<AnaliseInativacao> {
  const fd = basesForm(cadastro, estruturas);
  fd.append("itens", JSON.stringify(itens));
  fd.append("selecionados", JSON.stringify(selecionados));
  const res = await fetch("/api/inativacao/analisar", { method: "POST", body: fd });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new InativacaoApiError(data?.error || `Falha na requisição (${res.status})`, data?.code || "ERRO_INTERNO");
  }
  return data as AnaliseInativacao;
}

/** Efetiva a inativação: devolve o ZIP com a ficha e as estruturas atualizadas. */
export async function postExecutar(
  cadastro: File,
  estruturas: File,
  cpfs: string[],
  impressaoDigital: string,
  ignorarOrfas: boolean,
  onProgress?: (pct: number) => void,
): Promise<Blob> {
  const fd = basesForm(cadastro, estruturas);
  fd.append("cpfs", JSON.stringify(cpfs));
  fd.append("impressaoDigital", impressaoDigital);
  fd.append("ignore_orphan_warning", ignorarOrfas ? "true" : "false");
  try {
    return await postFormForBlob("/api/inativacao/executar", fd, onProgress);
  } catch (err) {
    if (err instanceof ApiRejection) throw new InativacaoApiError(err.message, err.code ?? "ERRO_INTERNO");
    throw err;
  }
}
```

- [ ] **Step 5: Run tests and typecheck**

Run: `npm --prefix frontend-react test -- src/lib/inativacaoApi.test.ts`
Expected: PASS (5 testes).

Run: `npm --prefix frontend-react exec tsc -- -b`
Expected: sem erros.

- [ ] **Step 6: Commit**

```bash
git add frontend-react/src/lib/api.ts frontend-react/src/lib/inativacaoApi.ts frontend-react/src/lib/inativacaoApi.test.ts
git commit -m "feat(frontend): cliente das rotas de análise e execução da inativação"
```

---

### Task 8: Hook `useInativacao` e regras do assistente (`wizard.ts`)

**Files:**
- Create: `frontend-react/src/tabs/InativacaoTab/wizard.ts`
- Modify: `frontend-react/src/tabs/InativacaoTab/useInativacao.ts` (o hook é reescrito; `classifyList`, `isValidFullName`, `formatCpf` e os tipos `Classification` **ficam**, com os testes atuais em `useInativacao.test.ts`)
- Test: `frontend-react/src/tabs/InativacaoTab/wizard.test.ts`, `frontend-react/src/tabs/InativacaoTab/useInativacao.flow.test.ts`

**Interfaces:**
- Consumes: Task 7 (`postAnalisar`, `postExecutar`, `InativacaoApiError`, tipos).
- Produces:
  - `wizard.ts`: `interface Confirmacao { impacto: boolean; orfas: boolean }`; `podeContinuar(analise: AnaliseInativacao | null): boolean`; `podeExecutar(analise, confirmacao): boolean`.
  - `useInativacao()` devolve: `listText, setListText, cadastro, setCadastro, estruturas, setEstruturas, classification, itens, podeAnalisar, analise, escolhidos, escolher(cpf, marcado), analisando, executando, progress, failure: GenerationFailure | null, concluido, analisar(): Promise<boolean>, executar(ignorarOrfas: boolean): Promise<boolean>, reset()`. Qualquer mudança em arquivos ou lista descarta a análise (`analise = null`, `escolhidos = []`, `failure = null`). `export const ARQUIVO_ZIP = "inativacao.zip"`.

- [ ] **Step 1: Write the failing tests — `wizard`**

`frontend-react/src/tabs/InativacaoTab/wizard.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import type { AnaliseInativacao, Situacao, UsuarioAnalise } from "../../lib/inativacaoApi";
import { podeContinuar, podeExecutar } from "./wizard";

function usuario(situacao: Situacao): UsuarioAnalise {
  return {
    cpf: "12345678909",
    cpfMascarado: "***.456.789-**",
    nome: "Ana Souza",
    email: "ana@x.com",
    situacao,
    alerta: null,
    estruturasViajante: [],
    comoAprovador: [],
    candidatos: [],
  };
}

function analise(usuarios: UsuarioAnalise[], orfas = 0): AnaliseInativacao {
  return {
    usuarios,
    resumo: {
      executaveis: usuarios.filter((u) => u.situacao === "EXECUTAVEL").length,
      estruturasExcluidas: 0,
      estruturasCompactadas: 0,
      estruturasOrfas: orfas,
      duplicados: [],
    },
    impressaoDigital: "d",
  };
}

describe("podeContinuar", () => {
  it("exige análise com ao menos um executável", () => {
    expect(podeContinuar(null)).toBe(false);
    expect(podeContinuar(analise([usuario("SEM_CPF")]))).toBe(false);
    expect(podeContinuar(analise([usuario("EXECUTAVEL")]))).toBe(true);
  });

  it("bloqueia enquanto houver homônimo sem escolha", () => {
    expect(podeContinuar(analise([usuario("EXECUTAVEL"), usuario("PENDENTE_SELECAO")]))).toBe(false);
  });
});

describe("podeExecutar", () => {
  it("exige a confirmação do impacto", () => {
    const a = analise([usuario("EXECUTAVEL")]);
    expect(podeExecutar(a, { impacto: false, orfas: false })).toBe(false);
    expect(podeExecutar(a, { impacto: true, orfas: false })).toBe(true);
  });

  it("com estrutura órfã exige também a ciência dela", () => {
    const a = analise([usuario("EXECUTAVEL")], 2);
    expect(podeExecutar(a, { impacto: true, orfas: false })).toBe(false);
    expect(podeExecutar(a, { impacto: true, orfas: true })).toBe(true);
  });

  it("sem análise nunca executa", () => {
    expect(podeExecutar(null, { impacto: true, orfas: true })).toBe(false);
  });
});
```

- [ ] **Step 2: Write the failing tests — hook**

`frontend-react/src/tabs/InativacaoTab/useInativacao.flow.test.ts`:

```ts
import { act, renderHook } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import * as api from "../../lib/inativacaoApi";
import type { AnaliseInativacao } from "../../lib/inativacaoApi";
import { useInativacao } from "./useInativacao";

vi.mock("../../lib/inativacaoApi", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/inativacaoApi")>()),
  postAnalisar: vi.fn(),
  postExecutar: vi.fn(),
}));
vi.mock("../../lib/downloadFile", () => ({ triggerAnchorDownload: vi.fn() }));
vi.mock("../../runs/runsStore", () => ({ addRun: vi.fn() }));
vi.mock("../../toast/toastStore", () => ({ pushToast: vi.fn() }));

const CPF = "12345678909";
const ANALISE: AnaliseInativacao = {
  usuarios: [
    {
      cpf: CPF,
      cpfMascarado: "***.456.789-**",
      nome: "Ana Souza",
      email: "ana@x.com",
      situacao: "EXECUTAVEL",
      alerta: null,
      estruturasViajante: [],
      comoAprovador: [],
      candidatos: [],
    },
    {
      cpf: null,
      cpfMascarado: "",
      nome: "Sem Cpf",
      email: "",
      situacao: "SEM_CPF",
      alerta: "sem cpf",
      estruturasViajante: [],
      comoAprovador: [],
      candidatos: [],
    },
  ],
  resumo: { executaveis: 1, estruturasExcluidas: 0, estruturasCompactadas: 0, estruturasOrfas: 0, duplicados: [] },
  impressaoDigital: "digital-1",
};

const cadastro = new File(["c"], "cadastro.xlsx");
const estruturas = new File(["e"], "estruturas.xlsx");

function preparado() {
  const hook = renderHook(() => useInativacao());
  act(() => {
    hook.result.current.setCadastro(cadastro);
    hook.result.current.setEstruturas(estruturas);
    hook.result.current.setListText(`${CPF}\nAna Souza`);
  });
  return hook;
}

beforeEach(() => {
  vi.mocked(api.postAnalisar).mockReset();
  vi.mocked(api.postExecutar).mockReset();
  Object.assign(URL, { createObjectURL: vi.fn(() => "blob:zip") });
});

describe("useInativacao", () => {
  it("só habilita a análise com as duas bases e ao menos um item", () => {
    const { result } = renderHook(() => useInativacao());
    expect(result.current.podeAnalisar).toBe(false);
    act(() => {
      result.current.setCadastro(cadastro);
      result.current.setEstruturas(estruturas);
      result.current.setListText(CPF);
    });
    expect(result.current.podeAnalisar).toBe(true);
  });

  it("analisa enviando as bases e a lista já classificada", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();

    let ok = false;
    await act(async () => {
      ok = await result.current.analisar();
    });

    expect(ok).toBe(true);
    expect(api.postAnalisar).toHaveBeenCalledWith(cadastro, estruturas, [CPF, "Ana Souza"], []);
    expect(result.current.analise).toEqual(ANALISE);
  });

  it("guarda a falha da análise com a mensagem do servidor", async () => {
    vi.mocked(api.postAnalisar).mockRejectedValue(new api.InativacaoApiError("Envie a base.", "BASE_AUSENTE"));
    const { result } = preparado();

    let ok = true;
    await act(async () => {
      ok = await result.current.analisar();
    });

    expect(ok).toBe(false);
    expect(result.current.analise).toBeNull();
    expect(result.current.failure?.message).toBe("Envie a base.");
  });

  it("mudar a lista, uma base ou as escolhas de outra análise descarta a análise", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });
    expect(result.current.analise).not.toBeNull();

    act(() => result.current.setListText(CPF));
    expect(result.current.analise).toBeNull();
  });

  it("executa só os CPFs executáveis, baixa o ZIP e marca como concluído", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    vi.mocked(api.postExecutar).mockResolvedValue(new Blob(["zip"]));
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    let ok = false;
    await act(async () => {
      ok = await result.current.executar(true);
    });

    expect(ok).toBe(true);
    expect(api.postExecutar).toHaveBeenCalledWith(cadastro, estruturas, [CPF], "digital-1", true, expect.any(Function));
    expect(result.current.concluido).toBe(true);
  });

  it("análise divergente no servidor descarta a análise e guarda a falha", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    vi.mocked(api.postExecutar).mockRejectedValue(new api.InativacaoApiError("A análise mudou.", "ANALISE_DIVERGENTE"));
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    await act(async () => {
      await result.current.executar(false);
    });

    expect(result.current.analise).toBeNull();
    expect(result.current.concluido).toBe(false);
    expect(result.current.failure?.message).toBe("A análise mudou.");
  });

  it("escolher marca e desmarca homônimos", () => {
    const { result } = renderHook(() => useInativacao());
    act(() => result.current.escolher(CPF, true));
    expect(result.current.escolhidos).toEqual([CPF]);
    act(() => result.current.escolher(CPF, false));
    expect(result.current.escolhidos).toEqual([]);
  });

  it("reset volta ao estado inicial", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });
    act(() => result.current.reset());
    expect(result.current.analise).toBeNull();
    expect(result.current.cadastro).toBeNull();
    expect(result.current.listText).toBe("");
  });
});
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `npm --prefix frontend-react test -- src/tabs/InativacaoTab`
Expected: FAIL — `./wizard` não existe e o hook antigo não tem `setCadastro`/`analisar`.

- [ ] **Step 4: Implement `wizard.ts`**

`frontend-react/src/tabs/InativacaoTab/wizard.ts`:

```ts
import type { AnaliseInativacao } from "../../lib/inativacaoApi";

export interface Confirmacao {
  impacto: boolean;
  orfas: boolean;
}

/** Do Impacto para o Confirmar: há quem inativar e nenhum homônimo ficou sem escolha. */
export function podeContinuar(analise: AnaliseInativacao | null): boolean {
  if (!analise) return false;
  return analise.resumo.executaveis > 0 && !analise.usuarios.some((u) => u.situacao === "PENDENTE_SELECAO");
}

/** Executar exige a confirmação do impacto e, havendo estrutura órfã, também a ciência dela. */
export function podeExecutar(analise: AnaliseInativacao | null, confirmacao: Confirmacao): boolean {
  if (!analise || !podeContinuar(analise)) return false;
  return confirmacao.impacto && (analise.resumo.estruturasOrfas === 0 || confirmacao.orfas);
}
```

- [ ] **Step 5: Rewrite the hook**

Em `frontend-react/src/tabs/InativacaoTab/useInativacao.ts`: **mantenha** de `EMAIL_PATTERN` até `formatCpf` (inclusive as interfaces `Classification`, `isValidFullName`, `classifyList`, `formatCpf`) e **apague** a interface `InativacaoResult` e a função `useInativacao`. Troque os imports do topo e acrescente o novo hook no fim:

```ts
import { useMemo, useState } from "react";
import { triggerAnchorDownload } from "../../lib/downloadFile";
import type { GenerationFailure } from "../../lib/failure";
import { InativacaoApiError, postAnalisar, postExecutar, type AnaliseInativacao } from "../../lib/inativacaoApi";
import { addRun } from "../../runs/runsStore";
import { pushToast } from "../../toast/toastStore";

export const ARQUIVO_ZIP = "inativacao.zip";
```

```ts
function comoFalha(err: unknown): GenerationFailure {
  return { message: err instanceof Error ? err.message : String(err) };
}

export function useInativacao() {
  const [listText, setListTextRaw] = useState("");
  const [cadastro, setCadastroRaw] = useState<File | null>(null);
  const [estruturas, setEstruturasRaw] = useState<File | null>(null);
  const [analise, setAnalise] = useState<AnaliseInativacao | null>(null);
  const [escolhidos, setEscolhidos] = useState<string[]>([]);
  const [analisando, setAnalisando] = useState(false);
  const [executando, setExecutando] = useState(false);
  const [progress, setProgress] = useState(0);
  const [failure, setFailure] = useState<GenerationFailure | null>(null);
  const [concluido, setConcluido] = useState(false);

  const classification = useMemo(() => classifyList(listText), [listText]);
  const itens = useMemo(
    () => [...classification.validCpfs, ...classification.validNames, ...classification.validEmails],
    [classification],
  );
  const podeAnalisar = Boolean(cadastro) && Boolean(estruturas) && itens.length > 0 && !analisando;

  /** Mexer nas entradas invalida a análise: o que o operador viu deixa de valer. */
  function invalidar() {
    setAnalise(null);
    setEscolhidos([]);
    setFailure(null);
  }

  function setListText(value: string) {
    setListTextRaw(value);
    invalidar();
  }

  function setCadastro(file: File | null) {
    setCadastroRaw(file);
    invalidar();
  }

  function setEstruturas(file: File | null) {
    setEstruturasRaw(file);
    invalidar();
  }

  function escolher(cpf: string, marcado: boolean) {
    setEscolhidos((atuais) => (marcado ? [...new Set([...atuais, cpf])] : atuais.filter((c) => c !== cpf)));
  }

  async function analisar(): Promise<boolean> {
    if (!cadastro || !estruturas) return false;
    setAnalisando(true);
    setFailure(null);
    try {
      setAnalise(await postAnalisar(cadastro, estruturas, itens, escolhidos));
      return true;
    } catch (err) {
      setAnalise(null);
      setFailure(comoFalha(err));
      return false;
    } finally {
      setAnalisando(false);
    }
  }

  async function executar(ignorarOrfas: boolean): Promise<boolean> {
    if (!cadastro || !estruturas || !analise) return false;
    const cpfs = analise.usuarios
      .filter((u) => u.situacao === "EXECUTAVEL" && u.cpf)
      .map((u) => u.cpf as string);
    setExecutando(true);
    setProgress(0);
    setFailure(null);
    try {
      const blob = await postExecutar(cadastro, estruturas, cpfs, analise.impressaoDigital, ignorarOrfas, setProgress);
      const url = URL.createObjectURL(blob);
      triggerAnchorDownload(url, ARQUIVO_ZIP);
      addRun({
        operation: "inativacao",
        inputSummary: [cadastro.name, estruturas.name, `${cpfs.length} usuário(s)`],
        outputFilename: ARQUIVO_ZIP,
        blobUrl: url,
      });
      pushToast("Inativação executada.", "success");
      setConcluido(true);
      return true;
    } catch (err) {
      // A análise que o operador viu não vale mais: volta a exigir uma nova.
      if (err instanceof InativacaoApiError && err.code === "ANALISE_DIVERGENTE") setAnalise(null);
      setFailure(comoFalha(err));
      return false;
    } finally {
      setExecutando(false);
      setProgress(0);
    }
  }

  function reset() {
    setListTextRaw("");
    setCadastroRaw(null);
    setEstruturasRaw(null);
    setAnalise(null);
    setEscolhidos([]);
    setFailure(null);
    setConcluido(false);
  }

  return {
    listText,
    setListText,
    cadastro,
    setCadastro,
    estruturas,
    setEstruturas,
    classification,
    itens,
    podeAnalisar,
    analise,
    escolhidos,
    escolher,
    analisando,
    executando,
    progress,
    failure,
    concluido,
    analisar,
    executar,
    reset,
  };
}
```

- [ ] **Step 6: Run tests and typecheck**

Run: `npm --prefix frontend-react test -- src/tabs/InativacaoTab`
Expected: PASS (`wizard.test.ts`, `useInativacao.flow.test.ts` e o `useInativacao.test.ts` atual).

`tsc -b` ainda **vai falhar** em `InativacaoTab/index.tsx` (usa o hook antigo): isso é esperado e a Task 9 resolve. Não rode o build completo agora.

- [ ] **Step 7: Commit**

```bash
git add frontend-react/src/tabs/InativacaoTab/wizard.ts frontend-react/src/tabs/InativacaoTab/wizard.test.ts frontend-react/src/tabs/InativacaoTab/useInativacao.ts frontend-react/src/tabs/InativacaoTab/useInativacao.flow.test.ts
git commit -m "feat(frontend): hook da inativação em duas etapas e regras do assistente"
```

---

### Task 9: Tela — `ImpactoCard` e o assistente de 3 etapas

**Files:**
- Create: `frontend-react/src/tabs/InativacaoTab/ImpactoCard.tsx`
- Test: `frontend-react/src/tabs/InativacaoTab/ImpactoCard.test.tsx`
- Rewrite: `frontend-react/src/tabs/InativacaoTab/index.tsx`

**Interfaces:**
- Consumes: Task 7 (tipos), Task 8 (`useInativacao`, `ARQUIVO_ZIP`, `podeContinuar`, `podeExecutar`, `Confirmacao`), componentes existentes `FileDropzone`, `Stepper`, `Card`, `Button`, `Field`, `Textarea`, `GenerationError`, `ProcessingProgress`, `PageHeader`, `Modal`, `RunHistoryPanel`, `explainFailure`.
- Produces: `ImpactoCard({ usuario, escolhidos, onEscolher })`; ids estáveis para o e2e: `#inativacao_cadastro`, `#inativacao_estruturas`, `#lista_text`, `#inativacao_btn` (analisar), `#inativacao_apply_selection_btn`, `#inativacao_next_btn`, `#inativacao_back_btn`, `#inativacao_confirm_impacto`, `#inativacao_confirm_orfas`, `#inativacao_execute_btn`, `#inativacao_new_btn`, `#inativacao_status`, `#inativacao_debug`.

- [ ] **Step 1: Write the failing test — `ImpactoCard`**

`frontend-react/src/tabs/InativacaoTab/ImpactoCard.test.tsx`:

```tsx
import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import type { UsuarioAnalise } from "../../lib/inativacaoApi";
import { ImpactoCard } from "./ImpactoCard";

function usuario(extra: Partial<UsuarioAnalise> = {}): UsuarioAnalise {
  return {
    cpf: "12345678909",
    cpfMascarado: "***.456.789-**",
    nome: "Ana Souza",
    email: "ana@x.com",
    situacao: "EXECUTAVEL",
    alerta: null,
    estruturasViajante: [],
    comoAprovador: [],
    candidatos: [],
    ...extra,
  };
}

function renderizar(u: UsuarioAnalise, escolhidos: string[] = [], onEscolher = vi.fn()) {
  render(
    <ul>
      <ImpactoCard usuario={u} escolhidos={escolhidos} onEscolher={onEscolher} />
    </ul>,
  );
  return onEscolher;
}

describe("ImpactoCard", () => {
  it("mostra o CPF mascarado e nunca o completo", () => {
    renderizar(usuario());
    expect(screen.getByText(/\*\*\*\.456\.789-\*\*/)).toBeInTheDocument();
    expect(screen.queryByText(/12345678909/)).not.toBeInTheDocument();
  });

  it("executável: mostra a estrutura do viajante a excluir e o impacto como aprovador", () => {
    renderizar(
      usuario({
        estruturasViajante: ["APR001"],
        comoAprovador: [
          { aprovacaoId: "APR010", posicoes: [2], segundoNivel: false, acao: "COMPACTACAO" },
          { aprovacaoId: "APR011", posicoes: [1], segundoNivel: true, acao: "ORFA" },
        ],
      }),
    );
    expect(screen.getByText(/Será excluída: APR001/)).toBeInTheDocument();
    expect(screen.getByText("Compactação")).toBeInTheDocument();
    expect(screen.getByText("Estrutura órfã")).toBeInTheDocument();
    expect(screen.getByText(/posição 2/)).toBeInTheDocument();
    expect(screen.getByText(/2º nível/)).toBeInTheDocument();
    expect(screen.getByText(/ficará sem nenhum aprovador/)).toBeInTheDocument();
  });

  it("executável sem estrutura nenhuma diz isso", () => {
    renderizar(usuario());
    expect(screen.getByText("Nenhuma estrutura direta encontrada.")).toBeInTheDocument();
    expect(screen.getByText("Não aparece como aprovador em outras estruturas.")).toBeInTheDocument();
  });

  it("sem CPF mostra o alerta e fica fora da cascata", () => {
    renderizar(usuario({ cpf: null, cpfMascarado: "", situacao: "SEM_CPF", alerta: "Sem CPF registrado." }));
    expect(screen.getByText("Sem CPF registrado.")).toBeInTheDocument();
    expect(screen.getByText("Sem CPF no cadastro")).toBeInTheDocument();
    expect(screen.queryByText("Compactação")).not.toBeInTheDocument();
  });

  it("homônimos: cada candidato é uma caixa de seleção; sem CPF não dá para escolher", async () => {
    const onEscolher = renderizar(
      usuario({
        cpf: null,
        situacao: "PENDENTE_SELECAO",
        candidatos: [
          { cpf: "11111111111", cpfMascarado: "***.111.111-**", nome: "João Silva", email: "j1@x.com" },
          { cpf: "", cpfMascarado: "", nome: "João Silva", email: "j2@x.com" },
        ],
      }),
    );
    const [comCpf, semCpf] = screen.getAllByRole("checkbox");
    expect(semCpf).toBeDisabled();
    await userEvent.click(comCpf);
    expect(onEscolher).toHaveBeenCalledWith("11111111111", true);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npm --prefix frontend-react test -- src/tabs/InativacaoTab/ImpactoCard.test.tsx`
Expected: FAIL — `./ImpactoCard` não existe.

- [ ] **Step 3: Implement `ImpactoCard.tsx`**

`frontend-react/src/tabs/InativacaoTab/ImpactoCard.tsx`:

```tsx
import type { EstruturaComoAprovador, UsuarioAnalise } from "../../lib/inativacaoApi";
import { cn } from "../../ui/cn";

const ROTULO: Record<UsuarioAnalise["situacao"], string> = {
  EXECUTAVEL: "Será inativado",
  SEM_CPF: "Sem CPF no cadastro",
  JA_INATIVO: "Já inativo",
  NAO_LOCALIZADO: "Não localizado",
  PENDENTE_SELECAO: "Escolha o usuário",
};

function descreverPosicao(e: EstruturaComoAprovador): string {
  const partes: string[] = [];
  if (e.posicoes.length === 1) partes.push(`posição ${e.posicoes[0]}`);
  if (e.posicoes.length > 1) partes.push(`posições ${e.posicoes.join(", ")}`);
  if (e.segundoNivel) partes.push("2º nível");
  return partes.join(" · ");
}

interface ImpactoCardProps {
  usuario: UsuarioAnalise;
  /** CPFs marcados entre os candidatos de homônimos. */
  escolhidos: string[];
  onEscolher: (cpf: string, marcado: boolean) => void;
}

/** Um usuário da análise: quem é, a situação e o que muda nas estruturas de aprovação. */
export function ImpactoCard({ usuario, escolhidos, onEscolher }: ImpactoCardProps) {
  const executavel = usuario.situacao === "EXECUTAVEL";
  const orfas = usuario.comoAprovador.filter((e) => e.acao === "ORFA").length;
  const identificacao = [usuario.cpfMascarado, usuario.email].filter(Boolean).join(" · ");

  return (
    <li
      data-situacao={usuario.situacao}
      className={cn(
        "rounded-control border px-4 py-3 text-sm",
        executavel ? "border-border bg-surface" : "border-warning/40 bg-warning-soft",
      )}
    >
      <div className="flex flex-wrap items-baseline justify-between gap-2">
        <p className="font-medium text-text">{usuario.nome || "—"}</p>
        <span
          className={cn(
            "rounded-pill border px-2 py-0.5 text-xs font-medium",
            executavel ? "border-success/40 bg-success-soft text-success" : "border-warning/40 text-warning",
          )}
        >
          {ROTULO[usuario.situacao]}
        </span>
      </div>
      {identificacao && <p className="mt-0.5 text-xs text-text-muted">{identificacao}</p>}
      {usuario.alerta && <p className="mt-2 text-warning">{usuario.alerta}</p>}

      {usuario.situacao === "PENDENTE_SELECAO" && (
        <fieldset className="mt-2">
          <legend className="text-text-muted">Há mais de um usuário com este nome. Marque quem deve ser inativado:</legend>
          {usuario.candidatos.map((c) => (
            <label key={`${c.cpf}-${c.email}`} className="mt-1.5 flex items-start gap-2 text-text">
              <input
                type="checkbox"
                className="mt-0.5 h-4 w-4"
                disabled={!c.cpf}
                checked={escolhidos.includes(c.cpf)}
                onChange={(e) => onEscolher(c.cpf, e.target.checked)}
              />
              <span>
                {c.nome} · {c.cpfMascarado || "sem CPF"} · {c.email || "sem e-mail"}
              </span>
            </label>
          ))}
        </fieldset>
      )}

      {executavel && (
        <dl className="mt-2 space-y-2">
          <div>
            <dt className="text-xs font-medium text-text-muted">Estrutura do viajante</dt>
            <dd className="text-text">
              {usuario.estruturasViajante.length > 0
                ? `Será excluída: ${usuario.estruturasViajante.join(", ")}`
                : "Nenhuma estrutura direta encontrada."}
            </dd>
          </div>
          <div>
            <dt className="text-xs font-medium text-text-muted">Como aprovador</dt>
            <dd className="text-text">
              {usuario.comoAprovador.length === 0 ? (
                "Não aparece como aprovador em outras estruturas."
              ) : (
                <ul className="space-y-1">
                  {usuario.comoAprovador.map((e) => (
                    <li key={e.aprovacaoId} className="flex flex-wrap items-center gap-2">
                      <span className="font-mono">{e.aprovacaoId}</span>
                      <span className="text-text-muted">{descreverPosicao(e)}</span>
                      <span
                        className={cn(
                          "rounded-pill border px-2 py-0.5 text-xs font-medium",
                          e.acao === "ORFA"
                            ? "border-danger/40 bg-danger-soft text-danger"
                            : "border-warning/40 bg-warning-soft text-warning",
                        )}
                      >
                        {e.acao === "ORFA" ? "Estrutura órfã" : "Compactação"}
                      </span>
                    </li>
                  ))}
                </ul>
              )}
              {orfas > 0 && (
                <p className="mt-1.5 text-danger">
                  Atenção: {orfas === 1 ? "1 estrutura ficará" : `${orfas} estruturas ficarão`} sem nenhum aprovador na
                  Argo.
                </p>
              )}
            </dd>
          </div>
        </dl>
      )}
    </li>
  );
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npm --prefix frontend-react test -- src/tabs/InativacaoTab/ImpactoCard.test.tsx`
Expected: PASS (5 testes). Se `toBeInTheDocument` não existir, confirme o `setupFiles` do vitest (`frontend-react/vite.config.ts`): os outros `*.test.tsx` do projeto já usam esses matchers.

- [ ] **Step 5: Rewrite the tab**

`frontend-react/src/tabs/InativacaoTab/index.tsx` (substitui o arquivo inteiro):

```tsx
import { useEffect, useRef, useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { explainFailure } from "../../lib/failure";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { Field } from "../../ui/Field";
import { GenerationError } from "../../ui/GenerationError";
import { IconCheck, IconUserMinus } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { ProcessingProgress } from "../../ui/ProcessingProgress";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { Stepper, type StepperItem } from "../../ui/Stepper";
import { Textarea } from "../../ui/Textarea";
import { ImpactoCard } from "./ImpactoCard";
import { ARQUIVO_ZIP, useInativacao } from "./useInativacao";
import { podeContinuar, podeExecutar, type Confirmacao } from "./wizard";

const ETAPAS = [
  {
    rotulo: "Bases e lista",
    titulo: "Enviar bases e lista",
    descricao: "Envie a base de cadastro e a de estruturas de aprovação e informe quem será inativado.",
  },
  {
    rotulo: "Impacto",
    titulo: "Conferir o impacto",
    descricao: "Veja o que muda antes de aplicar. Nada é alterado até você confirmar.",
  },
  {
    rotulo: "Confirmar",
    titulo: "Confirmar e executar",
    descricao: "Confirme para inativar os usuários e gerar os arquivos para carga.",
  },
];

const SEM_CONFIRMACAO = { digital: "", impacto: false, orfas: false };

function ArquivoEscolhido({ file, id, onClear }: { file: File | null; id: string; onClear: () => void }) {
  return (
    <div id={`${id}_feedback`} aria-live="polite" className="mt-3">
      {file && (
        <span className="inline-flex items-center gap-2 rounded-pill border border-border bg-surface-2 py-1 pl-3 pr-1.5 text-sm text-text">
          {file.name}
          <button
            id={`${id}_clear_btn`}
            type="button"
            title="Remover arquivo"
            aria-label={`Remover ${file.name}`}
            onClick={(e) => {
              e.stopPropagation();
              onClear();
            }}
            className="flex h-5 w-5 items-center justify-center rounded-pill text-text-subtle transition-colors hover:bg-danger/10 hover:text-danger"
          >
            ✕
          </button>
        </span>
      )}
    </div>
  );
}

function plural(n: number, singular: string, pluralForma: string): string {
  return `${n} ${n === 1 ? singular : pluralForma}`;
}

/**
 * Inativação em cascata como assistente de 3 etapas: bases e lista → impacto → confirmar.
 * A análise não altera nada; só a confirmação na última etapa dispara a execução.
 */
export function InativacaoTab() {
  const inativacao = useInativacao();
  const { analise, concluido, executando, analisando } = inativacao;
  const [etapa, setEtapa] = useState(0);
  const [helpOpen, setHelpOpen] = useState(false);
  const [confirmacao, setConfirmacao] = useState(SEM_CONFIRMACAO);

  const digital = analise?.impressaoDigital ?? "";
  // A confirmação vale só para a análise em que foi dada: outra análise, outra confirmação.
  const confirmada: Confirmacao = confirmacao.digital === digital ? confirmacao : SEM_CONFIRMACAO;
  // Sem análise válida (nada analisado ou ela foi descartada) só a primeira etapa faz sentido.
  const etapaAtual = analise || concluido ? etapa : 0;
  const resumo = analise?.resumo;
  const orfas = resumo?.estruturasOrfas ?? 0;
  const pendentes = analise ? analise.usuarios.filter((u) => u.situacao === "PENDENTE_SELECAO").length : 0;
  const travado = analisando || executando;

  function marcar(campo: keyof Confirmacao, valor: boolean) {
    setConfirmacao({ ...confirmada, digital, [campo]: valor });
  }

  async function handleAnalisar() {
    if (await inativacao.analisar()) setEtapa(1);
  }

  async function handleExecutar() {
    const ok = await inativacao.executar(orfas > 0 && confirmada.orfas);
    if (ok) setEtapa(2);
  }

  function novaInativacao() {
    inativacao.reset();
    setConfirmacao(SEM_CONFIRMACAO);
    setEtapa(0);
  }

  // Ao trocar de etapa o foco vai ao título: quem usa teclado ou leitor de tela começa a etapa do começo.
  const tituloRef = useRef<HTMLHeadingElement>(null);
  const primeiraRender = useRef(true);
  useEffect(() => {
    if (primeiraRender.current) {
      primeiraRender.current = false;
      return;
    }
    tituloRef.current?.focus({ preventScroll: true });
  }, [etapaAtual]);

  const detalhes = [
    inativacao.cadastro && inativacao.estruturas ? "2 bases enviadas" : "Cadastro e estruturas",
    resumo ? plural(resumo.executaveis, "a inativar", "a inativar") : "Conferir o impacto",
    concluido ? "ZIP gerado" : "Inativar e baixar",
  ];
  const passos: StepperItem[] = ETAPAS.map((e, i) => ({
    label: e.rotulo,
    detail: detalhes[i],
    state: concluido || i < etapaAtual ? "done" : i === etapaAtual ? "current" : "todo",
    selectable: !concluido && !travado && i < etapaAtual,
  }));

  const cabecalho = concluido
    ? { titulo: "Tudo pronto", descricao: "Comece outra inativação quando quiser." }
    : ETAPAS[etapaAtual];

  return (
    <div>
      <PageHeader
        title="Inativação"
        description="Inative usuários e mantenha as estruturas de aprovação consistentes."
        icon={<IconUserMinus className="h-5 w-5" />}
        actions={
          <Button
            id="inativacao_help_btn"
            variant="outline"
            size="sm"
            aria-label="Como usar a inativação"
            title="Guia da inativação"
            onClick={() => setHelpOpen(true)}
          >
            Como usar
          </Button>
        }
      />

      <Modal open={helpOpen} onClose={() => setHelpOpen(false)} title="Inativação">
        <ol className="list-decimal space-y-1 pl-5">
          <li>
            <strong>1º Passo:</strong> envie a <strong>base de cadastro</strong> e a <strong>base de estruturas</strong>{" "}
            (Excel) e cole <strong>CPFs</strong>, <strong>nomes completos</strong> ou <strong>e-mails</strong>, um por
            linha. Clique em <em>Analisar impacto</em>.
          </li>
          <li>
            <strong>2º Passo:</strong> confira, por usuário, a estrutura do viajante que será excluída e o efeito como
            aprovador (<em>compactação</em> ou <em>estrutura órfã</em>). Nada é alterado nesta etapa. Em caso de nomes
            repetidos, marque quem deve ser inativado.
          </li>
          <li>
            <strong>3º Passo:</strong> confirme e clique em <em>Executar inativação</em>. Você baixa um ZIP com a ficha de
            inativação e as estruturas atualizadas, prontos para a carga.
          </li>
        </ol>
        <p className="mt-2 text-text-subtle">
          Usuários sem CPF no cadastro não podem ser inativados por aqui: corrija o cadastro e analise de novo.
        </p>
      </Modal>

      <Stepper label="Etapas da inativação" steps={passos} onSelect={setEtapa} />

      <div className="grid items-start gap-6 xl:grid-cols-[minmax(0,1fr)_22rem]">
        <Card padding="lg">
          <header className="mb-4">
            <p className="text-xs font-medium text-text-muted">
              Etapa {etapaAtual + 1} de {ETAPAS.length}
            </p>
            <h3
              ref={tituloRef}
              id="inativacao_step_title"
              tabIndex={-1}
              className="text-lg font-semibold text-text outline-none"
            >
              {cabecalho.titulo}
            </h3>
            <p className="mt-0.5 text-sm text-text-muted">{cabecalho.descricao}</p>
          </header>

          {etapaAtual === 0 && (
            <>
              <div className="grid gap-4 md:grid-cols-2">
                <FileDropzone
                  id="inativacao_cadastro"
                  containerId="inativacao_cadastro_uploadArea"
                  accept=".xlsx,.xls"
                  ariaLabel="Upload da base de cadastro. Pressione para selecionar arquivo"
                  description="Base de cadastro de usuários (.xlsx)"
                  syncFiles={inativacao.cadastro ? [inativacao.cadastro] : []}
                  onFiles={(list) => inativacao.setCadastro(list[0] ?? null)}
                >
                  <ArquivoEscolhido
                    file={inativacao.cadastro}
                    id="inativacao_cadastro"
                    onClear={() => inativacao.setCadastro(null)}
                  />
                </FileDropzone>
                <FileDropzone
                  id="inativacao_estruturas"
                  containerId="inativacao_estruturas_uploadArea"
                  accept=".xlsx,.xls"
                  ariaLabel="Upload da base de estruturas de aprovação. Pressione para selecionar arquivo"
                  description="Base de estruturas de aprovação (.xlsx)"
                  syncFiles={inativacao.estruturas ? [inativacao.estruturas] : []}
                  onFiles={(list) => inativacao.setEstruturas(list[0] ?? null)}
                >
                  <ArquivoEscolhido
                    file={inativacao.estruturas}
                    id="inativacao_estruturas"
                    onClear={() => inativacao.setEstruturas(null)}
                  />
                </FileDropzone>
              </div>

              <div className="mt-4">
                <Field id="lista_text" label="Quem será inativado (um por linha)">
                  <Textarea
                    id="lista_text"
                    rows={4}
                    placeholder={"Ex: João Silva\n12345678901\nusuario@example.com"}
                    aria-describedby="lista_valid_summary lista_duplicates_warning"
                    value={inativacao.listText}
                    onChange={(e) => inativacao.setListText(e.target.value)}
                  />
                </Field>
                <div id="lista_valid_summary" className="mt-2 text-sm text-text-muted">
                  <span className="rounded bg-accent/10 px-2 py-0.5 font-medium text-accent-text">
                    {inativacao.classification.totalValid}
                  </span>{" "}
                  itens válidos (CPF, Nome Completo ou E-mail)
                </div>
                {inativacao.classification.duplicates.length > 0 && (
                  <div id="lista_duplicates_warning" className="mt-1 text-sm text-warning">
                    CPFs duplicados: {inativacao.classification.duplicates.join(", ")}
                  </div>
                )}
              </div>
            </>
          )}

          {etapaAtual === 1 && analise && resumo && (
            <>
              <p id="inativacao_summary" className="mb-3 text-sm text-text">
                {plural(resumo.executaveis, "usuário será inativado", "usuários serão inativados")} ·{" "}
                {plural(resumo.estruturasExcluidas, "estrutura excluída", "estruturas excluídas")} ·{" "}
                {plural(resumo.estruturasCompactadas, "compactada", "compactadas")} ·{" "}
                <span className={orfas > 0 ? "font-medium text-danger" : ""}>
                  {plural(orfas, "estrutura órfã", "estruturas órfãs")}
                </span>
              </p>
              <ul id="inativacao_impacto" className="max-h-[28rem] space-y-3 overflow-y-auto pr-1">
                {analise.usuarios.map((u, i) => (
                  <ImpactoCard
                    key={`${u.cpf ?? u.nome}-${i}`}
                    usuario={u}
                    escolhidos={inativacao.escolhidos}
                    onEscolher={inativacao.escolher}
                  />
                ))}
              </ul>
              {pendentes > 0 && (
                <div className="mt-3 flex flex-wrap items-center gap-3 text-sm text-warning">
                  <span>
                    {plural(pendentes, "nome repetido aguarda", "nomes repetidos aguardam")} sua escolha. Para deixar um
                    de fora, volte e tire-o da lista.
                  </span>
                  <Button
                    id="inativacao_apply_selection_btn"
                    variant="outline"
                    size="sm"
                    loading={analisando}
                    disabled={inativacao.escolhidos.length === 0}
                    onClick={() => void inativacao.analisar()}
                  >
                    Aplicar seleção
                  </Button>
                </div>
              )}
            </>
          )}

          {etapaAtual === 2 && analise && resumo && !concluido && (
            <>
              <p className="text-sm text-text">
                Serão inativados <strong>{plural(resumo.executaveis, "usuário", "usuários")}</strong>, com{" "}
                {plural(resumo.estruturasExcluidas, "estrutura excluída", "estruturas excluídas")} e{" "}
                {plural(resumo.estruturasCompactadas, "compactada", "compactadas")}
                {orfas > 0 && (
                  <>
                    , e <strong className="text-danger">{plural(orfas, "estrutura sem aprovador", "estruturas sem aprovador")}</strong>
                  </>
                )}
                .
              </p>
              <p className="mt-1 text-xs text-text-muted">
                Arquivos já carregados: {inativacao.cadastro?.name} e {inativacao.estruturas?.name}.
              </p>
              <div className="mt-4 space-y-2">
                <label className="flex items-start gap-2 text-sm text-text">
                  <input
                    id="inativacao_confirm_impacto"
                    type="checkbox"
                    className="mt-0.5 h-4 w-4"
                    checked={confirmada.impacto}
                    disabled={travado}
                    onChange={(e) => marcar("impacto", e.target.checked)}
                  />
                  <span>Revisei o impacto e quero inativar {plural(resumo.executaveis, "usuário", "usuários")}.</span>
                </label>
                {orfas > 0 && (
                  <label className="flex items-start gap-2 text-sm text-danger">
                    <input
                      id="inativacao_confirm_orfas"
                      type="checkbox"
                      className="mt-0.5 h-4 w-4"
                      checked={confirmada.orfas}
                      disabled={travado}
                      onChange={(e) => marcar("orfas", e.target.checked)}
                    />
                    <span>
                      Estou ciente de que {plural(orfas, "estrutura ficará", "estruturas ficarão")} sem nenhum aprovador
                      na Argo.
                    </span>
                  </label>
                )}
              </div>
              {executando && (
                <div className="mt-4">
                  <ProcessingProgress id="inativacao_progress" progress={inativacao.progress} />
                </div>
              )}
            </>
          )}

          {concluido && (
            <div
              id="inativacao_status"
              aria-live="polite"
              className="flex items-start gap-3 rounded-control border border-success/40 bg-success-soft px-4 py-3"
            >
              <span className="mt-0.5 text-success">
                <IconCheck className="h-5 w-5" />
              </span>
              <div className="text-sm">
                <p className="font-medium text-success">Inativação concluída</p>
                <p className="mt-0.5 text-text-muted">
                  O arquivo <strong className="font-medium text-text">{ARQUIVO_ZIP}</strong> foi gerado, com a ficha de
                  inativação e as estruturas atualizadas. Se o navegador não o salvou, baixe de novo em “Nesta sessão”.
                </p>
              </div>
            </div>
          )}

          {inativacao.failure && (
            <GenerationError
              id="inativacao_debug"
              className="mt-4"
              view={explainFailure(inativacao.failure, "Não foi possível concluir a inativação")}
            />
          )}

          <div className="mt-6 flex items-center justify-between gap-3 border-t border-border pt-4">
            {etapaAtual > 0 && !concluido ? (
              <Button id="inativacao_back_btn" variant="ghost" disabled={travado} onClick={() => setEtapa(etapaAtual - 1)}>
                Voltar
              </Button>
            ) : (
              <span />
            )}
            {concluido ? (
              <Button id="inativacao_new_btn" onClick={novaInativacao}>
                Nova inativação
              </Button>
            ) : etapaAtual === 0 ? (
              <Button
                id="inativacao_btn"
                aria-label="Analisar impacto da inativação"
                disabled={!inativacao.podeAnalisar}
                loading={analisando}
                onClick={handleAnalisar}
              >
                {analisando ? "Analisando..." : "Analisar impacto"}
              </Button>
            ) : etapaAtual === 1 ? (
              <Button
                id="inativacao_next_btn"
                title={pendentes > 0 ? "Escolha quem inativar entre os nomes repetidos" : undefined}
                disabled={!podeContinuar(analise)}
                onClick={() => setEtapa(2)}
              >
                Continuar
              </Button>
            ) : (
              <Button
                id="inativacao_execute_btn"
                variant={orfas > 0 ? "warning" : "primary"}
                disabled={!podeExecutar(analise, confirmada)}
                loading={executando}
                onClick={handleExecutar}
              >
                {executando ? "Executando..." : "Executar inativação"}
              </Button>
            )}
          </div>
        </Card>

        <div className="xl:sticky xl:top-20">
          <RunHistoryPanel operation="inativacao" className="mt-0" />
        </div>
      </div>
    </div>
  );
}
```

- [ ] **Step 6: Typecheck, lint and run the frontend suite**

Run: `npm --prefix frontend-react exec tsc -- -b`
Expected: sem erros. Se algum componente tiver props diferentes das usadas acima (`FileDropzone`, `ProcessingProgress`, `RunHistoryPanel`, `Button.variant`), ajuste **a chamada** ao contrato real do componente lendo o arquivo dele; não altere o componente compartilhado.

Run: `npm --prefix frontend-react run lint`
Expected: sem erros.

Run: `npm --prefix frontend-react test`
Expected: PASS (suíte inteira; os 205 testes anteriores continuam verdes).

- [ ] **Step 7: Commit**

```bash
git add frontend-react/src/tabs/InativacaoTab
git commit -m "feat(frontend): Inativação vira assistente de 3 etapas com painel de impacto"
```

---

### Task 10: Testes end-to-end (Playwright)

**Files:**
- Rewrite: `tests/e2e/inativacao-flow.spec.js`, `tests/e2e/preview-xss.spec.js`

**Interfaces:**
- Consumes: os ids da Task 9 e `xlsxFile`/`validCpf` de `tests/e2e/fixtures.mjs`.

- [ ] **Step 1: Reescreva `tests/e2e/inativacao-flow.spec.js`**

```js
// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const cadastro = () => [
  { CPF: validCpf(1), NomeCompleto: "Maria Silva", Email: "maria@x.com", Status: "ATIVO" },
  { CPF: validCpf(2), NomeCompleto: "Joao Pereira", Email: "joao@x.com", Status: "ATIVO" },
];

const estrutura = (id, cpfViajante, ...aprovadores) => {
  const linha = { AprovacaoId: id, AprovacaoPor: "VIAJANTE", CPF: cpfViajante, NomeViajante: `Viajante ${id}` };
  aprovadores.forEach((cpf, i) => {
    linha[`LoginAprovador_${i + 1}`] = cpf;
  });
  return linha;
};

/** S1: estrutura direta de Maria. S2: Maria aprova junto com Joao (compactação). */
const estruturas = () => [estrutura("S1", validCpf(1), validCpf(2)), estrutura("S2", validCpf(7), validCpf(2), validCpf(1))];

/** S3: Maria é a única aprovadora (estrutura órfã). */
const estruturasComOrfa = () => [estrutura("S3", validCpf(7), validCpf(1))];

async function analisar(page, bases, lista) {
  await page.setInputFiles("#inativacao_cadastro", xlsxFile("cadastro.xlsx", cadastro()));
  await page.setInputFiles("#inativacao_estruturas", xlsxFile("estruturas.xlsx", bases));
  await page.fill("#lista_text", lista);
  await page.locator("#inativacao_btn").click();
  await expect(page.locator("#inativacao_impacto li")).not.toHaveCount(0);
}

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#inativacao-tab").click();
  await expect(page.locator("#inativacao")).toBeVisible(); // painel da aba
  await expect(page.locator("#inativacao_btn")).toBeDisabled(); // sem bases nem lista
});

test("analisar mostra o impacto e executar baixa o ZIP", async ({ page }) => {
  await analisar(page, estruturas(), validCpf(1));

  const card = page.locator("#inativacao_impacto li").first();
  await expect(card).toContainText("Será excluída: S1");
  await expect(card).toContainText("Compactação");

  await page.locator("#inativacao_next_btn").click();
  await expect(page.locator("#inativacao_execute_btn")).toBeDisabled(); // falta a confirmação
  await page.locator("#inativacao_confirm_impacto").check();

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#inativacao_execute_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("inativacao.zip");
  await expect(page.locator("#inativacao_status")).toContainText("Inativação concluída");
});

test("estrutura órfã exige uma segunda confirmação", async ({ page }) => {
  await analisar(page, estruturasComOrfa(), validCpf(1));
  await expect(page.locator("#inativacao_impacto li").first()).toContainText("Estrutura órfã");

  await page.locator("#inativacao_next_btn").click();
  await page.locator("#inativacao_confirm_impacto").check();
  await expect(page.locator("#inativacao_execute_btn")).toBeDisabled(); // ainda falta a ciência da órfã
  await page.locator("#inativacao_confirm_orfas").check();
  await expect(page.locator("#inativacao_execute_btn")).toBeEnabled();
});

test("usuário sem CPF no cadastro não pode ser inativado", async ({ page }) => {
  const semCpf = [{ CPF: "", NomeCompleto: "Sem Cpf Silva", Email: "s@x.com", Status: "ATIVO" }];
  await page.setInputFiles("#inativacao_cadastro", xlsxFile("cadastro.xlsx", semCpf));
  await page.setInputFiles("#inativacao_estruturas", xlsxFile("estruturas.xlsx", estruturas()));
  await page.fill("#lista_text", "Sem Cpf Silva");
  await page.locator("#inativacao_btn").click();

  await expect(page.locator("#inativacao_impacto li").first()).toContainText("não possui CPF registrado");
  await expect(page.locator("#inativacao_next_btn")).toBeDisabled();
});

test("mexer na lista depois de analisar volta para a primeira etapa", async ({ page }) => {
  await analisar(page, estruturas(), validCpf(1));
  await page.locator("#inativacao_back_btn").click();
  await page.fill("#lista_text", validCpf(2));
  await expect(page.locator("#inativacao_btn")).toBeEnabled();
  await expect(page.locator("#inativacao_impacto")).toHaveCount(0);
});
```

- [ ] **Step 2: Reescreva `tests/e2e/preview-xss.spec.js`**

```js
// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const XSS = '<img src=x onerror="window.__xss=1"> "><script>window.__xss=1</script>';

test("nome malicioso da planilha aparece como texto na análise de inativacao", async ({ page }) => {
  let dialog = false;
  page.on("dialog", (d) => {
    dialog = true;
    d.dismiss().catch(() => {});
  });

  await page.goto("/");
  await page.locator("#inativacao-tab").click();

  const cadastro = [{ CPF: validCpf(1), NomeCompleto: XSS, Email: "a@x.com", Status: "ATIVO" }];
  const estruturas = [{ AprovacaoId: "S1", AprovacaoPor: "VIAJANTE", CPF: validCpf(1), LoginAprovador_1: validCpf(2) }];
  await page.setInputFiles("#inativacao_cadastro", xlsxFile("cadastro.xlsx", cadastro));
  await page.setInputFiles("#inativacao_estruturas", xlsxFile("estruturas.xlsx", estruturas));
  await page.fill("#lista_text", validCpf(1));
  await page.locator("#inativacao_btn").click();

  const card = page.locator("#inativacao_impacto li").first();
  await expect(card).toContainText("onerror"); // renderizado como texto
  expect(await card.evaluate((el) => el.querySelectorAll("img,script").length)).toBe(0);
  expect(dialog).toBe(false);
  expect(await page.evaluate(() => window.__xss)).toBeFalsy();
});
```

- [ ] **Step 3: Rode os e2e da Inativação**

Run: `npm run build:react && npx playwright test tests/e2e/inativacao-flow.spec.js tests/e2e/preview-xss.spec.js`
Expected: PASS. Se um seletor falhar, ajuste o **teste** ao id real da Task 9 (não o contrário); se o `FileDropzone` não expuser o `<input type=file>` pelo `id` informado, use o mesmo seletor que o e2e antigo usava (`#inativacao_base` era o `id` passado ao componente).

- [ ] **Step 4: Commit**

```bash
git add tests/e2e/inativacao-flow.spec.js tests/e2e/preview-xss.spec.js
git commit -m "test(e2e): fluxo de inativação em cascata e XSS na análise"
```

---

### Task 11: Documentação

**Files:**
- Modify: `REGRAS_APROVACAO_INATIVACAO.md`, `ARQUITETURA_MODERNIZADA.md`, `README.md`

- [ ] **Step 1: `REGRAS_APROVACAO_INATIVACAO.md`**

1. Na tabela do início, troque a linha da aba Inativação por:
   `| Inativação | \`backend/services/inactivation_cascade_service.py\` (\`InactivationCascadeService\`) + \`InactivationService\` + \`processar_inativacao_from_paths\` + \`ApprovalService\` + \`backend/api/inativacao.py\` | \`/api/inativacao/analisar\`, \`/api/inativacao/executar\` |`
2. **Apague** o bloco `> **As duas abas são independentes.** ...` e coloque no lugar:
   `> **Inativar um usuário atualiza as estruturas de aprovação.** A aba Inativação usa o \`ApprovalService\` para excluir a estrutura direta do viajante e compactar os aprovadores (§2). A aba Estruturas continua servindo para substituir/remover aprovadores sem inativar ninguém.`
3. Substitua **toda a seção "2. Inativação"** (do título `## 2. Inativação` até antes de `## 3. Comparativo`) por:

```markdown
## 2. Inativação em cascata

### 2.1 Objetivo

Inativar usuários da base do cliente e manter as estruturas de aprovação da
Argo consistentes: exclui a estrutura direta do viajante e remove o usuário de
todas as demais estruturas em que ele aprova, compactando as alçadas. Em duas
etapas: **análise** (`/api/inativacao/analisar`, sem efeito colateral) e
**execução** (`/api/inativacao/executar`, só depois da confirmação do operador).

### 2.2 Entradas

| Entrada | Conteúdo |
|---|---|
| Base de cadastro | usuários do cliente (CPF, nome, e-mail, status) |
| Base de estruturas | `AprovacaoId`, `AprovacaoPor`, **CPF do viajante**, `LoginAprovador_1..100`, opcional `LoginAprovador_SEGUNDO_NIVEL` |
| Lista | CPFs, nomes completos ou e-mails (até 500), colados, em planilha ou em JSON |

### 2.3 Situação de cada usuário

| Situação | Regra |
|---|---|
| `EXECUTAVEL` | encontrado, ATIVO, com CPF |
| `SEM_CPF` | encontrado sem CPF: **não executável**. Alerta: "Não foi possível mapear a Estrutura de Aprovação: Usuário encontrado no cadastro, mas não possui CPF registrado." |
| `JA_INATIVO` | status diferente de ATIVO no cadastro |
| `NAO_LOCALIZADO` | nenhum registro casou |
| `PENDENTE_SELECAO` | nome com mais de um registro: o operador escolhe quem inativar (só entra na cascata o escolhido) |

A busca é por CPF, e-mail e nome (`InactivationService.search_matches`) e o CPF
do cadastro é a chave para as estruturas (com o zero à esquerda restaurado).

### 2.4 Regras da cascata (ordem)

```
[1] Estrutura direta: linhas AprovacaoPor=VIAJANTE cujo CPF do viajante é o do
    usuário. São EXCLUÍDAS (Operacao=DELETE) e saem do cálculo seguinte.
[2] Compactação: nas demais estruturas, remove TODOS os CPFs da lista de
    LoginAprovador_1..100 e do 2º nível, compactando à esquerda; se o 1º nível
    esvazia e sobra um 2º nível que não está sendo removido, ele sobe.
[3] Órfã: estrutura que, depois de remover TODOS os CPFs da lista, fica sem
    nenhum aprovador. Estrutura excluída nunca conta como órfã.
```

Impacto por estrutura: `COMPACTACAO` (sobram aprovadores) ou `ORFA` (nenhum).
Estrutura órfã bloqueia a execução até o operador confirmar (`ignore_orphan_warning`).

### 2.5 Análise e execução

- `/analisar` devolve o diagnóstico e uma **impressão digital** (SHA-256 do
  impacto: CPFs executáveis, estruturas excluídas, compactadas e órfãs).
- `/executar` **recalcula** a análise a partir das planilhas e só prossegue se a
  impressão digital for a mesma que o operador viu (senão `409 ANALISE_DIVERGENTE`).
- Saída: ZIP `inativacao.zip` com `saida_inativacao.xlsx` (ficha `DELETE`) e
  `estruturas_atualizadas.xlsx` (todas as linhas das estruturas afetadas;
  `DELETE` nas excluídas, `UPDATE` nas alteradas). A base do cliente não é alterada.
- Auditoria: `inativacao_analise` e `inativacao_execucao`, só com contagens,
  impressão digital e **CPF mascarado** (`***.456.789-**`).

### 2.6 Erros

| HTTP | code | Quando |
|---|---|---|
| 400 | `BASE_AUSENTE` / `ARQUIVO_INVALIDO` / `BASE_SEM_COLUNA` | falta arquivo, arquivo ilegível ou coluna obrigatória ausente |
| 400 | `LISTA_VAZIA` / `LISTA_GRANDE` | lista vazia ou acima de 500 itens |
| 400 | `NADA_A_EXECUTAR` | nenhum CPF executável |
| 400 | `ORFAS_SEM_CONFIRMACAO` | há estruturas órfãs e a confirmação não veio |
| 409 | `ANALISE_DIVERGENTE` | a análise mudou desde a conferência |
| 413 | `ARQUIVO_GRANDE` | acima do teto de upload da rota (32 MB) |
| 500 | `ERRO_INTERNO` | erro inesperado (detalhe só no log) |
```

4. Na seção 3 (Comparativo), na coluna Inativação: `Entrada` → `Base de cadastro + base de estruturas + lista`; `Saída` → `ZIP: ficha DELETE + estruturas atualizadas`; `Efeito colateral automático no outro módulo` → `Atualiza estruturas (exclui a do viajante e compacta aprovadores)`; `Valida dígito verificador de CPF` → `Não (só comprimento)`.

- [ ] **Step 2: `ARQUITETURA_MODERNIZADA.md`**

1. Na tabela de endpoints, troque as 3 linhas de `/api/inativacao/buscar`, `/api/preview_inativacao` e `/api/process_inativacao` por:
   `| POST | \`/api/inativacao/analisar\` | análise de impacto da inativação (sem efeito colateral) |`
   `| POST | \`/api/inativacao/executar\` | executa em cascata e devolve o ZIP (ficha DELETE + estruturas atualizadas) |`
2. Troque o parágrafo `Removidos: ...` por:
   `Removidos: \`POST /api/inativacao/buscar\`, \`/api/preview_inativacao\` e \`/api/process_inativacao\` (substituídos pelo fluxo em duas etapas) e \`GET /health\` (duplicava \`/api/health\`). O antigo \`POST /api/inativacao/executar\` (sucesso falso, removido em 2026-09) foi reimplementado de verdade: lê as bases, gera os arquivos e audita.`
3. Em "Segurança", troque `o preview de inativação escapa` por `a análise de inativação exibe como texto (sem HTML)`; acrescente o item: `- **Privacidade**: a auditoria de inativação grava só CPF mascarado; nunca nome, e-mail nem CPF completo.`

- [ ] **Step 3: `README.md`**

1. Funcionalidades: troque o item **Inativação** por: `- **Inativação em cascata**: analisa o impacto (estrutura do viajante, aprovadores a compactar, estruturas que ficariam órfãs) e, confirmado, gera a ficha de desligamento (\`Operacao=DELETE\`) e as estruturas atualizadas num ZIP.`
2. Em "Inativação de usuários" (Fluxos principais), substitua os passos por:
   `1. Aba **Inativação** → envie a base de cadastro, a base de estruturas de aprovação e a lista (CPF, nome ou e-mail).`
   `2. Confira o impacto por usuário; escolha quem inativar em caso de nomes repetidos. Nada é alterado nesta etapa.`
   `3. Confirme (com ciência extra se houver estrutura órfã) e baixe o \`inativacao.zip\`.`
3. Na seção de testes, troque a linha `python -m pytest backend/tests/test_inativacao_api.py backend/tests/test_inativacao_buscar_api.py -q` por `python -m pytest backend/tests/test_inativacao_analisar_api.py backend/tests/test_inativacao_executar_api.py -q` e atualize o total da suíte (`123/123`) para o número real de `python -m pytest -q`.

- [ ] **Step 4: Commit**

```bash
git add REGRAS_APROVACAO_INATIVACAO.md ARQUITETURA_MODERNIZADA.md README.md
git commit -m "docs: inativação em cascata (regras, arquitetura e README)"
```

---

### Task 12: Verificação final e entrega

**Files:** nenhum novo.

- [ ] **Step 1: Suítes e linters completos**

```bash
python -m pytest -q
python -m ruff check backend
python -m ruff format --check backend
python -m mypy backend
npm --prefix frontend-react run lint
npm --prefix frontend-react test
npm --prefix frontend-react run build
npm run test:e2e
```
Expected: tudo verde. Anote as contagens (backend, Vitest, e2e) para o resumo.

- [ ] **Step 2: Conferência com a spec**

Releia `docs/superpowers/specs/2026-09-21-inativacao-cascata-design.md` e confirme, item a item: chave do viajante pela coluna CPF; ZIP com 2 planilhas (sempre as duas); `DELETE` na estrutura do viajante; 2º nível removido; aba substituída (rotas antigas devolvem 404); homônimos com seleção explícita; impressão digital (409); CPF mascarado na auditoria; teto de 500 itens e de 32 MB; erros com `code`. Qualquer divergência é bug do código **ou** do spec: corrija o que estiver errado e registre no commit.

- [ ] **Step 3: Rodar a aplicação e ver o fluxo funcionando**

Use a skill `run`: suba o servidor com o build novo e percorra, com planilhas de exemplo, análise → impacto → confirmação → download do ZIP; abra o ZIP e confira `Operacao` em cada planilha. Screenshots das 3 etapas devem mostrar o assistente sem quebras de layout (desktop e largura de celular).

- [ ] **Step 4: Finalização**

Use a skill `superpowers:finishing-a-development-branch`. **Não** faça push nem merge na `main` sem o usuário pedir.
