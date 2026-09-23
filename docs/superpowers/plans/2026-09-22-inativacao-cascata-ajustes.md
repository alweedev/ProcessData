# Inativação em Cascata — Ajustes Pós-Teste com Dados Reais Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Corrigir os gaps que apareceram ao testar a inativação em cascata com bases reais (Santillana) e ler o `Formulario Requisitos - Carga Aprovacao.pdf` da Argo — sem quebrar o formato aceito pela Argo para a carga real.

**Architecture:** Quatro ajustes independentes, todos no branch `feat/inativacao-cascata`:
1. A exigência de "CPF do viajante" na base de estruturas passa a ser condicional (só bloqueia quando existe alguma linha `AprovacaoPor=VIAJANTE`).
2. Quando não há coluna de CPF dedicada mas a base tem linhas VIAJANTE, o sistema detecta que `Valor` parece conter o CPF (mesmo padrão de login já usado na feature de substituir aprovador) e **pede confirmação explícita do operador** antes de usá-la — nunca assume sozinho.
3. Antes de liberar o ZIP para download, valida a base de estruturas exportada contra o schema oficial da Argo (`Formulario Requisitos - Carga Aprovacao.pdf`) e nunca exporta uma linha com `Operacao` vazio (campo obrigatório na carga).
4. A análise de impacto passa a devolver avisos de qualidade de dados (ex.: linhas VIAJANTE sem CPF reconhecível), no mesmo espírito do relatório de qualidade que já existe no Cadastro.

**Tech Stack:** Python 3.13 / Flask / Pandas (backend), React 19 + TypeScript + Vite (frontend), pytest, Vitest.

**Spec:** `docs/superpowers/specs/2026-09-21-inativacao-cascata-design.md` (spec original) + esta seção "Architecture" acima, que registra as 4 decisões tomadas em 2026-09-22 depois de testar com dados reais da Santillana e ler o PDF de requisitos da Argo.

## Global Constraints

- `Operacao` na base de estruturas exportada é campo obrigatório para a Argo: valores aceitos são só `INSERT`, `UPDATE` ou `DELETE` — nunca vazio (PDF, seção "Operacao").
- `LoginAprovador_N`: máximo de 12 caracteres cada (PDF).
- `AprovacaoPor`: só aceita `VIAJANTE, COMUNIDADE, MOTIVO, CCCLIENTE, CCEMPRESA, CLIENTE, EMPRESA, PROJETO` (PDF).
- `Tipo`: só aceita `U, S, P, N` (PDF).
- `Status` deve estar vazio na carga (PDF).
- "Login = CPF formatado `XXXXXXXXX-XX`" **não** é assumido automaticamente pelo sistema — exige confirmação explícita do operador quando inferido a partir de `Valor` (decisão do usuário em 2026-09-22).
- Toda mudança de código roda a suíte existente sem quebrar: `python -m pytest backend/tests/test_inactivation_cascade_service.py backend/tests/test_inativacao_analisar_api.py backend/tests/test_inativacao_executar_api.py -q` antes de cada commit.

---

## Task 1: Exigência de CPF do viajante vira condicional a existir VIAJANTE

**Files:**
- Modify: `backend/services/inactivation_cascade_service.py:65-76` (`_validar_estruturas`)
- Modify: `backend/services/inactivation_cascade_service.py:216-219` (chamada em `analisar`)
- Test: `backend/tests/test_inactivation_cascade_service.py`

**Interfaces:**
- Consumes: `ApprovalService.detect_approval_columns(df) -> dict` (já existe; chaves usadas: `aprovacao_id`, `aprovacao_por`, `approver_cols`, `traveler_cpf_col`).
- Produces: `_validar_estruturas(df_est: pd.DataFrame, cols: dict[str, Any]) -> None` (assinatura muda: ganha `df_est` como primeiro argumento; antes só recebia `cols`).

- [ ] **Step 1: Escrever o teste que falha — base sem nenhuma linha VIAJANTE não deve exigir CPF do viajante**

Adicionar em `backend/tests/test_inactivation_cascade_service.py` (junto dos outros testes de `analisar`):

```python
def test_base_sem_linha_viajante_nao_exige_cpf_do_viajante():
    df_cad = cad(USR_A)
    df_est = est({"AprovacaoId": "X1", "AprovacaoPor": "CCEMPRESA", "LoginAprovador_1": "999.999.999-99"})
    # Não deve levantar InativacaoError("BASE_SEM_COLUNA", ...): a base não usa estrutura por viajante.
    analise = InactivationCascadeService.analisar(df_cad, df_est, [A])
    assert analise.payload["usuarios"][0]["situacao"] == "EXECUTAVEL"
```

- [ ] **Step 2: Rodar o teste para confirmar que falha**

Run: `cd backend/.. && .venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py::test_base_sem_linha_viajante_nao_exige_cpf_do_viajante -v`
Expected: FAIL com `InativacaoError` / `BASE_SEM_COLUNA` (hoje bloqueia mesmo sem VIAJANTE).

- [ ] **Step 3: Implementar — validação condicional**

Em `backend/services/inactivation_cascade_service.py`, substituir `_validar_estruturas`:

```python
def _validar_estruturas(df_est: pd.DataFrame, cols: dict[str, Any]) -> None:
    faltando = []
    if not cols.get("aprovacao_id"):
        faltando.append("AprovacaoId")
    if not cols.get("aprovacao_por"):
        faltando.append("AprovacaoPor")
    if not cols.get("approver_cols"):
        faltando.append("LoginAprovador_1")
    if faltando:
        raise InativacaoError("BASE_SEM_COLUNA", "A base de estruturas não contém: " + ", ".join(faltando) + ".")

    por_col = cols["aprovacao_por"]
    tem_viajante = (df_est[por_col].astype(str).str.strip().str.upper() == "VIAJANTE").any()
    if tem_viajante and not cols.get("traveler_cpf_col"):
        raise InativacaoError("BASE_SEM_COLUNA", "A base de estruturas não contém: CPF (do viajante).")
```

E atualizar a chamada em `analisar` (linha ~219):

```python
        cols = ApprovalService.detect_approval_columns(df_est)
        _validar_estruturas(df_est, cols)
```

- [ ] **Step 4: Rodar o teste de novo e a suíte completa da cascata**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py -v`
Expected: todos PASS, incluindo o novo.

- [ ] **Step 5: Commit**

```bash
git add backend/services/inactivation_cascade_service.py backend/tests/test_inactivation_cascade_service.py
git commit -m "fix(backend): exige CPF do viajante só quando a base tem linhas VIAJANTE"
```

---

## Task 2: Detectar `Valor` como candidato de CPF do viajante, com confirmação explícita

**Files:**
- Modify: `backend/services/approval_service.py:367-424` (`detect_approval_columns`) e adicionar função nova
- Modify: `backend/services/inactivation_cascade_service.py` (`_validar_estruturas`, `analisar`, `executar`)
- Modify: `backend/api/inativacao.py` (rotas `/analisar` e `/executar`)
- Test: `backend/tests/test_approval_service.py` (ou arquivo de teste do `detect_approval_columns` já existente), `backend/tests/test_inactivation_cascade_service.py`, `backend/tests/test_inativacao_analisar_api.py`, `backend/tests/test_inativacao_executar_api.py`

**Interfaces:**
- Consumes: `_digits_matrix(df, columns) -> pd.DataFrame` (já existe em `approval_service.py:13`).
- Produces:
  - `ApprovalService.valor_parece_cpf_viajante(df: pd.DataFrame, cols: dict[str, Any]) -> bool` — novo método estático.
  - `InactivationCascadeService.analisar(df_cadastro, df_estruturas, itens, selecionados=None, confirmar_cpf_viajante=False) -> Analise` (novo parâmetro, default `False`).
  - `InactivationCascadeService.executar(df_cadastro, df_estruturas, cpfs, impressao_recebida, ignore_orphan_warning=False, confirmar_cpf_viajante=False) -> Execucao` (novo parâmetro).
  - Novo código de erro estável: `"CPF_VIAJANTE_A_CONFIRMAR"`, com `extra={"colunaCandidata": <nome da coluna>}`.

- [ ] **Step 1: Escrever o teste que falha — candidato detectado exige confirmação**

Em `backend/tests/test_inactivation_cascade_service.py`:

```python
def test_valor_com_cpf_em_linhas_viajante_exige_confirmacao_explicita():
    df_cad = cad(USR_A)
    df_est = est({
        "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": A,
        "LoginAprovador_1": "999.999.999-99",
    })
    with pytest.raises(InativacaoError) as exc:
        InactivationCascadeService.analisar(df_cad, df_est, [A])
    assert exc.value.code == "CPF_VIAJANTE_A_CONFIRMAR"
    assert exc.value.extra["colunaCandidata"] == "Valor"

    # Com confirmação explícita, passa a funcionar normalmente.
    analise = InactivationCascadeService.analisar(df_cad, df_est, [A], confirmar_cpf_viajante=True)
    assert analise.payload["usuarios"][0]["estruturasViajante"] == ["X1"]
```

- [ ] **Step 2: Rodar o teste para confirmar que falha**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py::test_valor_com_cpf_em_linhas_viajante_exige_confirmacao_explicita -v`
Expected: FAIL (hoje `BASE_SEM_COLUNA` sempre, sem candidato nem confirmação).

- [ ] **Step 3: Implementar `valor_parece_cpf_viajante` em `approval_service.py`**

Adicionar logo após `_digits_matrix` (linha ~28):

```python
def _valor_parece_cpf_viajante(df: pd.DataFrame, cols: dict[str, Any]) -> bool:
    """True se, nas linhas AprovacaoPor=VIAJANTE, a maioria dos valores de `Valor` tiver 11 dígitos.

    Detecta o padrão documentado pela Argo (Valor(VIAJANTE) = LOGIN, que neste sistema é o CPF
    formatado `XXXXXXXXX-XX`) quando não há coluna de CPF dedicada. Nunca usado sem confirmação
    explícita do operador — só sinaliza que dá para propor a confirmação.
    """
    valor_col = cols.get("valor")
    por_col = cols.get("aprovacao_por")
    if not valor_col or not por_col:
        return False
    is_traveler = df[por_col].astype(str).str.strip().str.upper() == "VIAJANTE"
    if not is_traveler.any():
        return False
    digits = _digits_matrix(df, [valor_col])[valor_col][is_traveler]
    validos = int((digits.str.len() == 11).sum())
    return validos > 0 and validos / len(digits) >= 0.8
```

E no `ApprovalService`, adicionar o método público (perto de `detect_approval_columns`):

```python
    @staticmethod
    def valor_parece_cpf_viajante(df: pd.DataFrame, cols: dict[str, Any]) -> bool:
        return _valor_parece_cpf_viajante(df, cols)
```

- [ ] **Step 4: Implementar a validação + confirmação em `inactivation_cascade_service.py`**

Substituir `_validar_estruturas` (já modificada na Task 1) para:

```python
def _validar_estruturas(df_est: pd.DataFrame, cols: dict[str, Any], confirmar_cpf_viajante: bool) -> None:
    faltando = []
    if not cols.get("aprovacao_id"):
        faltando.append("AprovacaoId")
    if not cols.get("aprovacao_por"):
        faltando.append("AprovacaoPor")
    if not cols.get("approver_cols"):
        faltando.append("LoginAprovador_1")
    if faltando:
        raise InativacaoError("BASE_SEM_COLUNA", "A base de estruturas não contém: " + ", ".join(faltando) + ".")

    por_col = cols["aprovacao_por"]
    tem_viajante = (df_est[por_col].astype(str).str.strip().str.upper() == "VIAJANTE").any()
    if not tem_viajante or cols.get("traveler_cpf_col"):
        return
    if not ApprovalService.valor_parece_cpf_viajante(df_est, cols):
        raise InativacaoError("BASE_SEM_COLUNA", "A base de estruturas não contém: CPF (do viajante).")
    if not confirmar_cpf_viajante:
        raise InativacaoError(
            "CPF_VIAJANTE_A_CONFIRMAR",
            "Não há coluna de CPF do viajante na base, mas a coluna 'Valor' parece conter o CPF nas linhas "
            "VIAJANTE. Confirme para usá-la nesta análise.",
            extra={"colunaCandidata": cols.get("valor")},
        )
```

E em `analisar` (assinatura e corpo):

```python
    @staticmethod
    def analisar(
        df_cadastro: pd.DataFrame,
        df_estruturas: pd.DataFrame,
        itens: list[str],
        selecionados: list[str] | None = None,
        confirmar_cpf_viajante: bool = False,
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
        _validar_estruturas(df_est, cols, confirmar_cpf_viajante)
        if confirmar_cpf_viajante and not cols.get("traveler_cpf_col"):
            cols["traveler_cpf_col"] = cols.get("valor")
```

(o restante do corpo de `analisar` continua igual — `find_traveler_structures` e `structures_with_foreign_rows` já filtram por `AprovacaoPor=VIAJANTE` antes de ler `cols["traveler_cpf_col"]`, então funcionam sem mudança apontando para `Valor`.)

E em `executar`, propagar o parâmetro na assinatura e na chamada interna a `analisar`:

```python
    @staticmethod
    def executar(
        df_cadastro: pd.DataFrame,
        df_estruturas: pd.DataFrame,
        cpfs: list[str],
        impressao_recebida: str,
        ignore_orphan_warning: bool = False,
        confirmar_cpf_viajante: bool = False,
    ) -> Execucao:
        lista = sorted({c for c in (clean_cpf(x) for x in (cpfs or [])) if c})
        if not lista:
            raise InativacaoError("NADA_A_EXECUTAR", "Nenhum usuário foi informado para inativar.")
        analise = InactivationCascadeService.analisar(
            df_cadastro, df_estruturas, lista, selecionados=lista, confirmar_cpf_viajante=confirmar_cpf_viajante
        )
```

(o resto do corpo de `executar` não muda.)

- [ ] **Step 5: Propagar o parâmetro nas rotas Flask**

Em `backend/api/inativacao.py`, em `api_inativacao_analisar`:

```python
        itens = _extrair_itens(paths)
        confirmar_cpf_viajante = request.form.get("confirmarCpfViajante", "").lower() in _VERDADEIRO
        analise = InactivationCascadeService.analisar(
            df_cadastro, df_estruturas, itens, _json_lista("selecionados"), confirmar_cpf_viajante
        )
```

E em `api_inativacao_executar`:

```python
        cpfs = _json_lista("cpfs")
        ignorar_orfas = request.form.get("ignore_orphan_warning", "").lower() in _VERDADEIRO
        confirmar_cpf_viajante = request.form.get("confirmarCpfViajante", "").lower() in _VERDADEIRO
        execucao = InactivationCascadeService.executar(
            df_cadastro,
            df_estruturas,
            cpfs,
            request.form.get("impressaoDigital", ""),
            ignorar_orfas,
            confirmar_cpf_viajante,
        )
```

- [ ] **Step 6: Rodar os testes**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py backend/tests/test_inativacao_analisar_api.py backend/tests/test_inativacao_executar_api.py -v`
Expected: todos PASS.

- [ ] **Step 7: Commit**

```bash
git add backend/services/approval_service.py backend/services/inactivation_cascade_service.py backend/api/inativacao.py backend/tests/test_inactivation_cascade_service.py
git commit -m "feat(backend): detecta CPF do viajante em Valor e exige confirmação explícita do operador"
```

---

## Task 3: Nunca exportar `Operacao` vazio + validar schema Argo antes do download

**Files:**
- Create: `backend/services/argo_schema_validator.py`
- Test: `backend/tests/test_argo_schema_validator.py`
- Modify: `backend/services/inactivation_cascade_service.py:executar` (filtrar linhas sem `Operacao` e validar antes de retornar)
- Modify: `backend/api/inativacao.py:api_inativacao_executar` (tratar novo código de erro)

**Interfaces:**
- Consumes: nenhuma nova (só `pandas.DataFrame`).
- Produces: `ArgoSchemaValidator.validar(df: pd.DataFrame) -> list[str]` — lista de violações em texto (vazia = OK). Usado por `InactivationCascadeService.executar`, que levanta `InativacaoError("SCHEMA_ARGO_INVALIDO", ...)` se a lista não for vazia.

- [ ] **Step 1: Escrever o teste que falha — validador rejeita `Operacao` vazio e login grande demais**

Criar `backend/tests/test_argo_schema_validator.py`:

```python
import pandas as pd

from backend.services.argo_schema_validator import ArgoSchemaValidator


def test_aceita_planilha_valida():
    df = pd.DataFrame([
        {"Operacao": "DELETE", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-11", "Status": ""},
    ])
    assert ArgoSchemaValidator.validar(df) == []


def test_rejeita_operacao_vazia():
    df = pd.DataFrame([
        {"Operacao": "", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-11", "Status": ""},
    ])
    erros = ArgoSchemaValidator.validar(df)
    assert any("Operacao" in e for e in erros)


def test_rejeita_login_aprovador_maior_que_12_caracteres():
    df = pd.DataFrame([
        {"Operacao": "UPDATE", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-111-extra", "Status": ""},
    ])
    erros = ArgoSchemaValidator.validar(df)
    assert any("LoginAprovador_1" in e for e in erros)


def test_rejeita_status_preenchido():
    df = pd.DataFrame([
        {"Operacao": "DELETE", "AprovacaoId": "X1", "AprovacaoPor": "VIAJANTE", "Valor": "12345678901",
         "Tipo": "S", "LoginAprovador_1": "111111111-11", "Status": "Ativo"},
    ])
    erros = ArgoSchemaValidator.validar(df)
    assert any("Status" in e for e in erros)
```

- [ ] **Step 2: Rodar para confirmar que falha**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_argo_schema_validator.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'backend.services.argo_schema_validator'`.

- [ ] **Step 3: Implementar o validador**

Criar `backend/services/argo_schema_validator.py`:

```python
"""Valida a base de estruturas exportada contra o schema oficial da carga de aprovação da Argo.

Regras de `Formulario Requisitos - Carga Aprovacao.pdf`: campos obrigatórios, valores aceitos e
tamanhos máximos. Roda só antes do download (nunca bloqueia a análise) — se falhar aqui é sinal de
bug na geração, não de dado ruim do cliente.
"""

import re

import pandas as pd

_OPERACAO_VALIDA = {"INSERT", "UPDATE", "DELETE"}
_APROVACAO_POR_VALIDA = {
    "VIAJANTE", "COMUNIDADE", "MOTIVO", "CCCLIENTE", "CCEMPRESA", "CLIENTE", "EMPRESA", "PROJETO",
}
_TIPO_VALIDO = {"U", "S", "P", "N"}
_LOGIN_MAX = 12


class ArgoSchemaValidator:
    @staticmethod
    def validar(df: pd.DataFrame) -> list[str]:
        erros: list[str] = []
        if df.empty:
            return erros

        if "Operacao" not in df.columns or (df["Operacao"].astype(str).str.strip() == "").any():
            erros.append("Operacao: há linha(s) sem valor (obrigatório: INSERT, UPDATE ou DELETE).")
        elif not df["Operacao"].astype(str).str.strip().str.upper().isin(_OPERACAO_VALIDA).all():
            erros.append("Operacao: há valor fora de INSERT, UPDATE ou DELETE.")

        if "AprovacaoPor" in df.columns:
            preenchidos = df["AprovacaoPor"].astype(str).str.strip()
            invalidos = preenchidos[(preenchidos != "") & ~preenchidos.str.upper().isin(_APROVACAO_POR_VALIDA)]
            if not invalidos.empty:
                erros.append("AprovacaoPor: há valor fora da lista aceita pela Argo.")

        if "Tipo" in df.columns:
            preenchidos = df["Tipo"].astype(str).str.strip()
            invalidos = preenchidos[(preenchidos != "") & ~preenchidos.str.upper().isin(_TIPO_VALIDO)]
            if not invalidos.empty:
                erros.append("Tipo: há valor fora de U, S, P ou N.")

        for col in df.columns:
            if re.match(r"(?i)^LoginAprovador_\d+$", str(col)):
                grandes = df[col].astype(str).str.len() > _LOGIN_MAX
                if grandes.any():
                    erros.append(f"{col}: há login com mais de {_LOGIN_MAX} caracteres.")

        if "Status" in df.columns and (df["Status"].astype(str).str.strip() != "").any():
            erros.append("Status: deve vir vazio na carga.")

        return erros
```

- [ ] **Step 4: Rodar o teste do validador**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_argo_schema_validator.py -v`
Expected: todos PASS.

- [ ] **Step 5: Ligar o validador em `executar` e nunca exportar `Operacao` vazio**

Em `backend/services/inactivation_cascade_service.py`, importar no topo:

```python
from backend.services.argo_schema_validator import ArgoSchemaValidator
```

E em `executar`, depois de montar `estruturas` (antes do `resumo`/`return`):

```python
        estruturas = estruturas[estruturas["Operacao"].astype(str).str.strip() != ""].reset_index(drop=True)

        erros_schema = ArgoSchemaValidator.validar(estruturas)
        if erros_schema:
            raise InativacaoError(
                "SCHEMA_ARGO_INVALIDO",
                "A exportação ficou fora do formato aceito pela Argo: " + "; ".join(erros_schema),
                status=500,
                extra={"erros": erros_schema},
            )
```

(mantém o resto do corpo igual; `resumo["linhasEstruturas"]` passa a refletir só as linhas que de fato mudaram, já que as `Operacao=""` saem do dataframe.)

- [ ] **Step 6: Escrever teste de integração — execução não inclui linha sem mudança real**

Em `backend/tests/test_inactivation_cascade_service.py`, localizar um teste de `executar` existente com uma estrutura compactada onde nem toda posição muda (se não houver, adicionar):

```python
def test_executar_nunca_inclui_linha_de_estrutura_sem_operacao():
    df_cad = cad(USR_A)
    df_est = est(viajante("X1", A, B))  # A é o viajante, B aprovador em outra estrutura não entra aqui
    analise = InactivationCascadeService.analisar(df_cad, df_est, [A])
    execucao = InactivationCascadeService.executar(df_cad, df_est, [A], analise.payload["impressaoDigital"])
    assert (execucao.estruturas["Operacao"].astype(str).str.strip() != "").all()
```

- [ ] **Step 7: Rodar toda a suíte da cascata**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py backend/tests/test_argo_schema_validator.py backend/tests/test_inativacao_executar_api.py -v`
Expected: todos PASS.

- [ ] **Step 8: Commit**

```bash
git add backend/services/argo_schema_validator.py backend/services/inactivation_cascade_service.py backend/tests/test_argo_schema_validator.py backend/tests/test_inactivation_cascade_service.py
git commit -m "feat(backend): valida schema da Argo antes do download e nunca exporta Operacao vazio"
```

---

## Task 4: Avisos de qualidade de dados na análise (backend)

**Files:**
- Modify: `backend/services/inactivation_cascade_service.py` (`Analise.payload`, dentro de `analisar`)
- Test: `backend/tests/test_inactivation_cascade_service.py`

**Interfaces:**
- Produces: `analise.payload["avisos"] -> list[str]` (nova chave no JSON de `/api/inativacao/analisar`, sempre presente, lista vazia = sem avisos).

- [ ] **Step 1: Escrever o teste que falha**

```python
def test_avisos_sinalizam_linhas_viajante_sem_cpf_reconhecivel():
    df_cad = cad(USR_A)
    df_est = est(
        viajante("X1", A),
        {"AprovacaoId": "X2", "AprovacaoPor": "VIAJANTE", "CPF": "", "LoginAprovador_1": "999.999.999-99"},
    )
    analise = InactivationCascadeService.analisar(df_cad, df_est, [A])
    assert any("sem CPF" in a for a in analise.payload["avisos"])
```

- [ ] **Step 2: Rodar para confirmar que falha**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py::test_avisos_sinalizam_linhas_viajante_sem_cpf_reconhecivel -v`
Expected: FAIL com `KeyError: 'avisos'`.

- [ ] **Step 3: Implementar — calcular avisos em `analisar`**

Em `inactivation_cascade_service.py`, dentro de `analisar`, antes do `payload = {...}` (por volta da linha 262):

```python
        avisos: list[str] = []
        if cols.get("traveler_cpf_col"):
            por_col = cols["aprovacao_por"]
            is_traveler = df_est[por_col].astype(str).str.strip().str.upper() == "VIAJANTE"
            cpf_col = cols["traveler_cpf_col"]
            sem_cpf = int((is_traveler & (df_est[cpf_col].apply(clean_cpf).str.len() != 11)).sum())
            if sem_cpf:
                avisos.append(f"{sem_cpf} estrutura(s) VIAJANTE sem CPF reconhecível na base (ficam fora da análise).")
        if busca.get("duplicates"):
            avisos.append(f"{len(busca['duplicates'])} CPF(s) duplicado(s) na lista informada.")
```

E incluir no `payload`:

```python
        payload = {
            "usuarios": usuarios,
            "resumo": {...},  # inalterado
            "impressaoDigital": digital,
            "avisos": avisos,
        }
```

- [ ] **Step 4: Rodar os testes**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inactivation_cascade_service.py -v`
Expected: todos PASS.

- [ ] **Step 5: Ajustar `_detalhes_analise` em `backend/api/inativacao.py` para não vazar nada sensível dos avisos na auditoria**

Os avisos aqui só citam contagens, não CPF/nome — conferir em `backend/api/inativacao.py:100-108` que `_detalhes_analise` inclui `resumo` (que já não tem dado pessoal). Adicionar `avisos` ao que vai para a auditoria, já que também é só contagem:

```python
def _detalhes_analise(analise: Analise) -> dict:
    """Só contagens, impressão digital e CPF mascarado: nome e e-mail nunca vão para o histórico."""
    resumo = dict(analise.payload["resumo"])
    resumo["duplicados"] = [mascarar_cpf(c) for c in resumo.get("duplicados", [])]
    return {
        "resumo": resumo,
        "avisos": analise.payload.get("avisos", []),
        "impressaoDigital": analise.payload["impressaoDigital"],
        "usuarios": [{"cpf": u["cpfMascarado"], "situacao": u["situacao"]} for u in analise.payload["usuarios"]],
    }
```

- [ ] **Step 6: Rodar a suíte da API de análise**

Run: `.venv/Scripts/python.exe -m pytest backend/tests/test_inativacao_analisar_api.py backend/tests/test_inactivation_cascade_service.py -v`
Expected: todos PASS.

- [ ] **Step 7: Commit**

```bash
git add backend/services/inactivation_cascade_service.py backend/api/inativacao.py backend/tests/test_inactivation_cascade_service.py
git commit -m "feat(backend): avisos de qualidade de dados na análise de impacto da cascata"
```

---

## Task 5: Frontend — cliente da API e hook propagam confirmação e avisos

**Files:**
- Modify: `frontend-react/src/lib/inativacaoApi.ts`
- Modify: `frontend-react/src/tabs/InativacaoTab/useInativacao.ts`
- Test: `frontend-react/src/tabs/InativacaoTab/useInativacao.test.ts`

**Interfaces:**
- Consumes: rotas backend da Task 2 (`confirmarCpfViajante` em `FormData`) e Task 4 (`avisos` na resposta).
- Produces:
  - `InativacaoApiError` ganha campo `extra?: Record<string, unknown>`.
  - `AnaliseInativacao` ganha campo `avisos: string[]`.
  - `postAnalisar(cadastro, estruturas, itens, selecionados, confirmarCpfViajante: boolean)`.
  - `postExecutar(cadastro, estruturas, cpfs, impressaoDigital, ignorarOrfas, confirmarCpfViajante: boolean, onProgress?)`.
  - `useInativacao()` ganha `cpfViajanteCandidato: string | null` e `confirmarCpfViajante(): void`.

- [ ] **Step 1: Escrever o teste que falha**

Em `frontend-react/src/tabs/InativacaoTab/useInativacao.test.ts`, adicionar (seguindo o padrão de mock já usado no arquivo para `postAnalisar`):

```typescript
it("expõe cpfViajanteCandidato quando o servidor pede confirmação e reanalisa ao confirmar", async () => {
  const erro = new InativacaoApiError("Confirme para usar Valor.", "CPF_VIAJANTE_A_CONFIRMAR");
  erro.extra = { colunaCandidata: "Valor" };
  vi.mocked(postAnalisar).mockRejectedValueOnce(erro).mockResolvedValueOnce(ANALISE_OK);

  const { result } = renderHook(() => useInativacao());
  act(() => {
    result.current.setCadastro(new File([""], "cadastro.xlsx"));
    result.current.setEstruturas(new File([""], "estruturas.xlsx"));
    result.current.setListText("12345678901");
  });

  await act(async () => {
    await result.current.analisar();
  });
  expect(result.current.cpfViajanteCandidato).toBe("Valor");

  await act(async () => {
    await result.current.confirmarCpfViajanteEAnalisar();
  });
  expect(result.current.cpfViajanteCandidato).toBeNull();
  expect(result.current.analise).toEqual(ANALISE_OK);
});
```

- [ ] **Step 2: Rodar para confirmar que falha**

Run: `cd frontend-react && npm run test -- useInativacao.test.ts`
Expected: FAIL (`confirmarCpfViajanteEAnalisar is not a function` / `cpfViajanteCandidato` undefined).

- [ ] **Step 3: Implementar — `inativacaoApi.ts`**

```typescript
export class InativacaoApiError extends Error {
  readonly code: string;
  extra?: Record<string, unknown>;

  constructor(message: string, code: string, extra?: Record<string, unknown>) {
    super(message);
    this.name = "InativacaoApiError";
    this.code = code;
    this.extra = extra;
  }
}
```

Em `postAnalisar`, capturar `data?.colunaCandidata` e outros campos de `extra` no erro, e adicionar o parâmetro:

```typescript
export async function postAnalisar(
  cadastro: File,
  estruturas: File,
  itens: string[],
  selecionados: string[],
  confirmarCpfViajante: boolean,
): Promise<AnaliseInativacao> {
  const fd = basesForm(cadastro, estruturas);
  fd.append("itens", JSON.stringify(itens));
  fd.append("selecionados", JSON.stringify(selecionados));
  fd.append("confirmarCpfViajante", confirmarCpfViajante ? "true" : "false");
  const res = await fetch("/api/inativacao/analisar", { method: "POST", body: fd });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new InativacaoApiError(data?.error || `Falha na requisição (${res.status})`, data?.code || "ERRO_INTERNO", {
      colunaCandidata: data?.colunaCandidata,
    });
  }
  if (!Array.isArray(data?.usuarios)) throw new InativacaoApiError("Resposta inesperada do servidor.", "ERRO_INTERNO");
  return { ...data, avisos: Array.isArray(data?.avisos) ? data.avisos : [] } as AnaliseInativacao;
}
```

Atualizar `AnaliseInativacao`:

```typescript
export interface AnaliseInativacao {
  usuarios: UsuarioAnalise[];
  resumo: ResumoAnalise;
  impressaoDigital: string;
  avisos: string[];
}
```

E `postExecutar` ganha o mesmo parâmetro, repassado no `FormData` como `confirmarCpfViajante`.

- [ ] **Step 4: Implementar — `useInativacao.ts`**

Adicionar estado e função:

```typescript
  const [cpfViajanteCandidato, setCpfViajanteCandidato] = useState<string | null>(null);
```

Em `invalidar()`, resetar: `setCpfViajanteCandidato(null);`.

Em `analisar`, aceitar a confirmação corrente e tratar o código específico:

```typescript
  async function analisar(confirmarCpfViajante = false): Promise<boolean> {
    if (!cadastro || !estruturas || analisandoRef.current) return false;
    const minha = geracao.current;
    analisandoRef.current = true;
    setAnalisando(true);
    setFailure(null);
    try {
      const resultado = await postAnalisar(cadastro, estruturas, itens, escolhidos, confirmarCpfViajante);
      if (minha !== geracao.current) return false;
      setAnalise(resultado);
      setCpfViajanteCandidato(null);
      return true;
    } catch (err) {
      if (minha !== geracao.current) return false;
      setAnalise(null);
      if (err instanceof InativacaoApiError && err.code === "CPF_VIAJANTE_A_CONFIRMAR") {
        setCpfViajanteCandidato((err.extra?.colunaCandidata as string) || "Valor");
        return false;
      }
      setFailure(comoFalha(err));
      return false;
    } finally {
      if (minha === geracao.current) {
        analisandoRef.current = false;
        setAnalisando(false);
      }
    }
  }

  async function confirmarCpfViajanteEAnalisar(): Promise<boolean> {
    return analisar(true);
  }
```

`executar` passa a enviar `Boolean(cpfViajanteCandidato === null && analise !== null)`... na verdade mais simples: guardar a confirmação usada com sucesso em um ref e reenviá-la:

```typescript
  const confirmarCpfViajanteRef = useRef(false);
```

Em `analisar`, no `try` bem-sucedido, `confirmarCpfViajanteRef.current = confirmarCpfViajante;`. Em `executar`, passar `confirmarCpfViajanteRef.current` para `postExecutar`.

Exportar `cpfViajanteCandidato` e `confirmarCpfViajanteEAnalisar` no `return` do hook.

- [ ] **Step 5: Rodar os testes do hook**

Run: `cd frontend-react && npm run test -- useInativacao.test.ts`
Expected: todos PASS.

- [ ] **Step 6: Commit**

```bash
git add frontend-react/src/lib/inativacaoApi.ts frontend-react/src/tabs/InativacaoTab/useInativacao.ts frontend-react/src/tabs/InativacaoTab/useInativacao.test.ts
git commit -m "feat(frontend): hook e client propagam confirmação de CPF do viajante e avisos"
```

---

## Task 6: Frontend — UI de confirmação e painel de avisos

**Files:**
- Modify: `frontend-react/src/tabs/InativacaoTab/index.tsx`
- Test: `frontend-react/src/tabs/InativacaoTab/index.test.tsx`

**Interfaces:**
- Consumes: `useInativacao()` (Task 5): `cpfViajanteCandidato`, `confirmarCpfViajanteEAnalisar`, `analise.avisos`.

- [ ] **Step 1: Escrever o teste que falha**

Em `frontend-react/src/tabs/InativacaoTab/index.test.tsx`, seguindo o padrão de mock de `useInativacao` já usado no arquivo, adicionar um caso onde `cpfViajanteCandidato` vem preenchido e afirmar que o banner e o botão de confirmação aparecem, e que clicar chama `confirmarCpfViajanteEAnalisar`.

- [ ] **Step 2: Rodar para confirmar que falha**

Run: `cd frontend-react && npm run test -- InativacaoTab`
Expected: FAIL (elemento do banner não existe ainda).

- [ ] **Step 3: Implementar o banner de confirmação**

Em `InativacaoTab/index.tsx`, dentro do bloco `{etapaAtual === 0 && (...)}`, depois dos avisos de duplicados (linha ~277), adicionar:

```tsx
                {inativacao.cpfViajanteCandidato && (
                  <div
                    id="inativacao_cpf_viajante_confirm"
                    className="mt-3 rounded-control border border-warning/40 bg-warning-soft px-4 py-3 text-sm"
                  >
                    <p className="text-text">
                      Não encontramos uma coluna de CPF do viajante dedicada na base de estruturas, mas a coluna{" "}
                      <strong>{inativacao.cpfViajanteCandidato}</strong> parece conter o CPF nas linhas VIAJANTE.
                    </p>
                    <Button
                      id="inativacao_confirm_cpf_viajante_btn"
                      variant="outline"
                      size="sm"
                      className="mt-2"
                      loading={inativacao.analisando}
                      onClick={() => void inativacao.confirmarCpfViajanteEAnalisar()}
                    >
                      Usar {inativacao.cpfViajanteCandidato} como CPF do viajante e analisar
                    </Button>
                  </div>
                )}
```

- [ ] **Step 4: Implementar o painel de avisos na etapa de impacto**

No bloco `{etapaAtual === 1 && analise && resumo && (...)}`, logo depois do `<p id="inativacao_summary">` (linha ~291), adicionar:

```tsx
              {analise.avisos.length > 0 && (
                <ul id="inativacao_avisos" className="mb-3 space-y-1 text-sm text-warning">
                  {analise.avisos.map((aviso, i) => (
                    <li key={i}>⚠ {aviso}</li>
                  ))}
                </ul>
              )}
```

- [ ] **Step 5: Rodar os testes do componente**

Run: `cd frontend-react && npm run test -- InativacaoTab`
Expected: todos PASS.

- [ ] **Step 6: Rodar toda a suíte de unidade do frontend e o typecheck**

Run: `cd frontend-react && npm run test && npm run build`
Expected: testes PASS, build sem erro de tipo.

- [ ] **Step 7: Commit**

```bash
git add frontend-react/src/tabs/InativacaoTab/index.tsx frontend-react/src/tabs/InativacaoTab/index.test.tsx
git commit -m "feat(frontend): UI de confirmação do CPF do viajante e avisos de qualidade no impacto"
```

---

## Verificação final

- [ ] Rodar a suíte inteira do backend: `.venv/Scripts/python.exe -m pytest -q`
- [ ] Rodar a suíte inteira do frontend: `cd frontend-react && npm run test`
- [ ] Rebuildar o frontend e testar manualmente com os arquivos reais da Santillana (`Mapa de Usuarios - Santillana.xlsx` como cadastro, `Carga aprovacao - Santillana.xls` **original, sem a coluna extra `CPFViajante`** como estruturas): confirmar que aparece o banner de confirmação, que confirmar destrava a análise, que os avisos aparecem na etapa de impacto, e que o ZIP final passa pelo `ArgoSchemaValidator` sem erro.
- [ ] Apagar o arquivo de workaround gerado nesta conversa (`Carga aprovacao - Santillana (com CPFViajante).xlsx`) — não deve mais ser necessário.
