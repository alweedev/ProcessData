# Regras de Negócio — Aprovação e Inativação

Cobre as duas abas que operam sobre a base de usuários/estruturas já
cadastrada na Argo (diferente do cadastro, que **cria** registros — ver
`REGRAS_NEGOCIO_PROCESSAMENTO.md`).

| Aba | Código | Endpoints |
|---|---|---|
| Estruturas de Aprovação | `backend/services/approval_service.py` (`ApprovalService`) + `backend/api/aprovacao.py` | `/api/aprovacao/remover/{preview,export}`, `/api/aprovacao/substituir/{preview,export}` |
| Inativação | `backend/services/inactivation_cascade_service.py` (`InactivationCascadeService`) + `InactivationService` + `processar_inativacao_from_paths` + `ApprovalService` + `backend/api/inativacao.py` | `/api/inativacao/analisar`, `/api/inativacao/executar` |

> **Inativar um usuário atualiza as estruturas de aprovação.** A aba Inativação usa o `ApprovalService` para excluir a estrutura direta do viajante e compactar os aprovadores (§2). A aba Estruturas continua servindo para substituir/remover aprovadores sem inativar ninguém.

---

## 1. Estruturas de Aprovação

### 1.1 Conceitos-chave

- Uma **estrutura** é uma linha (ou grupo de linhas) da base de carga,
  identificada por `AprovacaoId`. Pode ter vários aprovadores em
  `LoginAprovador_1` até `LoginAprovador_100`, mais um opcional
  `LoginAprovador_SEGUNDO_NIVEL`.
- `AprovacaoPor` indica se a aprovação é por **VIAJANTE** (pessoa específica)
  ou **CCEMPRESA** (centro de custo).
- Colunas são detectadas automaticamente e case-insensitive
  (`ApprovalService.detect_approval_columns`): `AprovacaoId`,
  `LoginAprovador_1..100`, e opcionalmente `AprovacaoPor`, `Aprovacao`,
  `Tipo`, `Valor`, `DescricaoCCusto`, `CodigoCCusto`,
  `LoginAprovador_SEGUNDO_NIVEL`, `SegundoNivelMaster`, `NomeViajante`.

### 1.2 Validação de CPF e base de usuários

```
✓ CPF informado, 11 dígitos, dígito verificador válido (módulo 11)
✓ CPF existe na base de usuários
✓ Status = ATIVO na base de usuários (quando a base tem coluna Status)

✗ "Informe um CPF para o aprovador."
✗ "CPF inválido. Informe 11 dígitos."
✗ "CPF inválido (dígito verificador)."
✗ "CPF não encontrado na base de usuários."
✗ "Aprovador não está ATIVO na base de usuários (Status: '...')."
```

> A checagem de **dígito verificador** só existe aqui (`is_valid_cpf`,
> módulo 11); cadastro valida só o comprimento (com a ressalva do zero à
> esquerda — ver `REGRAS_NEGOCIO_PROCESSAMENTO.md` §6).
> A base de usuários precisa ter coluna `CPF` e uma de nome (`NomeCompleto`
> ou `Nome` [+ `SobreNome`]) — senão: `"Base de usuários não contém coluna
> 'CPF'."` / `"...coluna de nome ('NomeCompleto' ou 'Nome')."`.
>
> **Na prática, `"CPF inválido. Informe 11 dígitos."` só aparece com CPF
> longo demais** (12+ dígitos): `normalize_cpf_input` também faz
> `zfill(11)` antes de checar o tamanho, então uma entrada curta (ex.:
> `"123456"`) vira 11 dígitos e cai direto no erro de dígito verificador —
> que, ao contrário do cadastro, sempre pega esses casos porque a checagem
> de módulo 11 é muito mais rígida que a de comprimento.

### 1.3 Substituição de aprovador (`ApprovalService.replace_cpf`)

Cenário real: um aprovador sai e outro assume o lugar dele nas mesmas
estruturas. **Não existe** um fluxo de "inserir aprovador do zero em
estruturas arbitrárias" — decidir em quais estruturas e em qual posição
inserir não tem regra de negócio definida hoje.

```
[1] Validação: os 2 CPFs (11 díg. + dígito verificador + ATIVO na base),
    CPF novo != CPF atual, base tem AprovacaoId + LoginAprovador_1..100

[2] Detecção: localiza estruturas onde o CPF atual aparece
    (LoginAprovador_1..100 e, opcionalmente, SEGUNDO_NIVEL)

[3] Duplicidade: verifica se o CPF novo já é aprovador em alguma dessas
    estruturas — no nível principal OU no segundo nível
    (estruturasComDuplicidade)

[4] Preview: estruturas afetadas, posições, quais já teriam duplicidade

[5] Confirmação: usuário decide se ignora a duplicidade
    (ignore_duplicate_warning=true) ou cancela

[6] Execução: substitui o CPF atual pelo novo, na MESMA posição, sem
    compactar — a estrutura não muda de forma, só o login muda.
    Se replace_second_level=true e o CPF atual está no SEGUNDO_NIVEL,
    substitui lá também (padrão: false, segundo nível intocado).
    Sem gate automático de deduplicação: se o CPF novo já estava em outra
    posição da mesma estrutura e o aviso foi ignorado, ele fica duplicado
    de propósito — a decisão é do usuário.

[7] Export: exporta TODAS as linhas das estruturas alvo (a Argo precisa da
    estrutura completa); Operacao="UPDATE" só nas linhas alteradas.
```

### 1.4 Remoção de aprovador (`ApprovalService.remove_cpf_and_compact`)

```
[1] Preview: estruturas afetadas, posições, quais ficarão sem NENHUM
    aprovador, agrupadas por AprovacaoPor

[2] Confirmação do usuário

[3] Execução: remove o CPF de LoginAprovador_1..100 e COMPACTA à esquerda
    (sem posições vazias no meio).
    Se o 1º nível ficou vazio e há SEGUNDO_NIVEL preenchido com outro CPF:
    promove o 2º nível para o 1º e esvazia o 2º nível.
    Se remove_second_level=true e o 2º nível é o CPF removido: apaga o 2º
    nível também.
    Gate: se alguma estrutura ficar sem NENHUM aprovador, retorna aviso
    (HTTP 400 + warning), só prossegue com ignore_empty_warning=true.

[4] Export: TODAS as linhas das estruturas alvo; Operacao="UPDATE" só nas
    linhas alteradas.
```

### 1.5 Exportação parcial (`mode=selected`)

Os dois fluxos aceitam `mode=all` (padrão) ou `mode=selected` +
`selected_aprovacao_ids[]` — a UI pré-seleciona todas as estruturas do
preview e permite desmarcar. A checagem de duplicidade/estrutura-vazia roda
só sobre o conjunto alvo (`target_ids`), não sobre todas as estruturas
afetadas pelo CPF.

### 1.6 Formato de CPF exibido

`format_cpf_for_output()` produz **`XXXXXXXXX-XX`** (9 dígitos + traço + 2
dígitos verificadores, **sem pontos**) — mesmo formato usado no Login do
cadastro. Exemplo de resposta de preview (substituição):

```jsonc
{
  "oldApprover": { "cpf": "123456789-00", "nomeCompleto": "João Silva" },
  "newApprover": { "cpf": "987654321-00", "nomeCompleto": "Maria Souza" },
  "summary": {
    "estruturasAfetadas": 25, "ocorrenciasTotal": 42,
    "porAprovacaoPor": { "VIAJANTE": 15, "CCEMPRESA": 10 },
    "estruturasComDuplicidade": 2
  },
  "items": [
    {
      "aprovacaoId": "APR001", "aprovacaoPor": "VIAJANTE",
      "posicoes": [1, 5], "segundoNivel": false, "teraDuplicidade": false
    }
  ]
}
```

### 1.7 Casos de erro

| Situação | Erro |
|---|---|
| CPF vazio | `Informe um CPF para o aprovador.` |
| CPF inválido (tamanho) | `CPF inválido. Informe 11 dígitos.` |
| CPF inválido (dígito verificador) | `CPF inválido (dígito verificador).` |
| CPF não existe na base de usuários | `CPF não encontrado na base de usuários.` |
| CPF existe mas não está ATIVO | `Aprovador não está ATIVO na base de usuários (Status: '...').` |
| Base sem coluna CPF | `Base de usuários não contém coluna 'CPF'.` |
| Base sem coluna de nome | `Base de usuários não contém coluna de nome ('NomeCompleto' ou 'Nome').` |
| CPF novo == CPF atual (só substituição) | `O CPF do novo aprovador deve ser diferente do CPF atual.` |
| CPF não aparece em nenhuma estrutura | `CPF não está presente em nenhuma estrutura de aprovação.` |

---

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
| `EXECUTAVEL` | encontrado, com CPF e ATIVO (ver abaixo) |
| `SEM_CPF` | encontrado sem CPF: **não executável**. Alerta: "Não foi possível mapear a Estrutura de Aprovação: Usuário encontrado no cadastro, mas não possui CPF registrado." |
| `JA_INATIVO` | cadastro com coluna de status e status diferente de ATIVO. Alerta: "Usuário não está ATIVO no cadastro (Status: '...')." |
| `NAO_LOCALIZADO` | nenhum registro casou |
| `PENDENTE_SELECAO` | nome digitado que consta em mais de um registro do cadastro: o operador escolhe quem inativar |

- **Status**: com coluna de status no cadastro, só `Status == "ATIVO"` é
  executável; **qualquer outro valor, inclusive vazio**, é `JA_INATIVO`. Sem
  coluna de status, os usuários encontrados são executáveis.
- **Homônimos**: quando o nome digitado consta em mais de uma linha do cadastro,
  os registros casados por nome exigem escolha explícita, mesmo que outro
  homônimo tenha sido digitado por CPF ou e-mail (esse continua executável).
  Sem escolha, nada é processado para os ambíguos. Na tela, o botão "Continuar"
  fica bloqueado enquanto houver homônimo pendente; para deixar um de fora, o
  operador remove o nome da lista.

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
- Na aba, a confirmação do operador vale só para a análise em que foi dada
  (amarrada à impressão digital): refazer a análise exige confirmar de novo. Se
  houver estrutura órfã, é preciso uma **segunda confirmação**, específica.
- Saída: ZIP `inativacao.zip` com `saida_inativacao.xlsx` (ficha `DELETE`) e
  `estruturas_atualizadas.xlsx` (todas as linhas das estruturas afetadas;
  `DELETE` nas excluídas, `UPDATE` nas alteradas). A base do cliente não é alterada.
  Um CPF com várias linhas ATIVO no cadastro gera uma linha da ficha por linha.
- Resumo da execução: `usuariosInativados` conta **usuários (CPFs)**, não linhas
  da ficha; por isso pode ser menor que o número de linhas de `saida_inativacao.xlsx`.
- Auditoria: `inativacao_analise` e `inativacao_execucao` registram só contagens,
  a impressão digital, a `situacao` de cada usuário e **CPFs mascarados** (`***.456.789-**`, também os CPFs
  duplicados na lista). Nunca gravam nome, e-mail nem CPF completo.

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

As rotas antigas `/api/inativacao/buscar`, `/api/process_inativacao` e
`/api/preview_inativacao` deixaram de existir.

---

## 3. Comparativo

| Aspecto | Aprovação | Inativação |
|---|---|---|
| Ação | Substitui/remove aprovadores em estruturas existentes | Inativa usuários (ficha DELETE) e atualiza as estruturas de aprovação em cascata |
| Entrada | Base de estruturas + base de usuários + CPF(s) | Base de cadastro + base de estruturas + lista |
| Chave de match | `AprovacaoId` / CPF do aprovador | CPF do cadastro (a busca aceita CPF, e-mail ou nome; tudo é resolvido para o CPF do cadastro) |
| Valida dígito verificador de CPF | Sim | Não (só comprimento) |
| Saída | Base de estruturas atualizada (`Operacao=UPDATE` nas linhas alteradas) | ZIP: ficha DELETE + estruturas atualizadas |
| Efeito colateral automático no outro módulo | Nenhum | Atualiza estruturas (exclui a do viajante e compacta aprovadores) |
