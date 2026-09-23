# Inativação em cascata (usuário + estruturas de aprovação) — design

## Contexto e objetivo

Hoje a aba **Inativação** só gera a ficha `Operacao=DELETE` a partir de uma base
de usuários. Inativar alguém **não** limpa as estruturas de aprovação da Argo:
sobram a estrutura direta do viajante e o nome dele como aprovador em outras
estruturas, e o operador precisa lembrar de fazer isso à mão na aba Estruturas
(`REGRAS_APROVACAO_INATIVACAO.md`, nota inicial: "as duas abas são
independentes").

Objetivo: a inativação passa a ser uma operação em **duas etapas** — análise de
impacto sem efeito colateral e execução confirmada — que inativa o usuário e
mantém as estruturas de aprovação consistentes, reaproveitando o
`ApprovalService` em vez de duplicar regra.

Requisitos não funcionais (fazem parte do aceite, não são detalhe):
correção com dados pessoais (CPF), desempenho previsível com bases grandes,
erros claros e testáveis, e nada de estado escondido no servidor.

## Decisões fechadas

| Tema | Decisão |
|---|---|
| Chave da estrutura do viajante | Linhas `AprovacaoPor=VIAJANTE` cuja coluna **CPF** da própria base de estruturas é o CPF do usuário |
| Escopo | **Lista** de usuários (como a aba tem hoje), não um por vez |
| Saída da execução | **ZIP com 2 planilhas**: `saida_inativacao.xlsx` (ficha `DELETE`) e `estruturas_atualizadas.xlsx`. A base do cliente **não** é alterada |
| Estrutura do viajante na saída | Mantida no arquivo com `Operacao=DELETE` |
| 2º nível | O usuário inativado é removido também de `LoginAprovador_SEGUNDO_NIVEL`, com a promoção ao 1º nível já existente |
| Aba atual | **Substituída** pelo novo fluxo; as duas bases passam a ser obrigatórias |
| Homônimos (busca por nome) | Seleção explícita do operador; nada é processado sem escolha |
| Arquitetura | Sem estado: as duas rotas recebem as planilhas; a execução **recalcula** a análise no servidor |

## Estado atual que este design respeita

- `ApprovalService.remove_cpf_and_compact` já remove um CPF, compacta
  `LoginAprovador_1..100` e promove o 2º nível; `_check_structures_without_approvers`
  já detecta estrutura que ficaria sem aprovador. É o núcleo dos cenários A/B.
- `InactivationService.search_matches` já localiza por CPF, e-mail e nome;
  `process_from_dataframes` já gera a ficha `DELETE` (só usuários ATIVO).
- `POST /api/inativacao/executar` **existiu e foi removida em 2026-09**: devolvia
  sucesso sem ler a base, gerar arquivo nem auditar. A nova rota tem de ser
  testada contra exatamente esse defeito (ver Testes).
- Uploads são apagados no `finally` de cada requisição; `MAX_CONTENT_LENGTH`
  é 16 MB, global (`backend/core/config.py`, aplicado em `backend/app.py`).
- `AuditService.record` grava o dicionário `details` como vier; **não há
  máscara de CPF** no projeto.

## Arquitetura

```
api/inativacao.py            rotas /analisar e /executar (finas)
services/
  inactivation_cascade_service.py   orquestra a cascata (regra de negócio nova)
  approval_service.py               estendido (métodos abaixo), sem duplicar
  inactivation_service.py           busca e ficha DELETE (reuso)
  audit_service.py                  eventos + máscara de CPF
shared/
  cpf_mask.py                       mascarar_cpf("12345678901") -> "***.456.789-**"
  fingerprint.py                    impressão digital estável do diagnóstico
```

**`ApprovalService` — novos métodos (todos puros, sem I/O, sem mutar entrada):**

- `find_traveler_structures(df, cpf_digits, cols)` — ids das estruturas
  `AprovacaoPor=VIAJANTE` do usuário.
- `delete_structures(df, ids)` — as linhas dessas estruturas com `Operacao=DELETE`.
- `remove_cpfs_and_compact(df, cpfs, cols, target_ids, remove_second_level)` —
  **uma passada** sobre a base contra o *conjunto* de CPFs.
  `remove_cpf_and_compact` passa a delegar a este método com um conjunto de
  1 CPF: uma só implementação, e a aba Estruturas herda o ganho de desempenho.
  A suíte `test_aprovacao*.py` existente é a rede de segurança dessa mudança.
- `structures_left_without_approvers(df, cpfs, cols, target_ids, remove_second_level)`
  — generalização do `_check_structures_without_approvers` para conjunto de CPFs.

**`InactivationCascadeService`** só orquestra: resolve a lista em usuários,
chama os métodos acima na ordem correta e monta o diagnóstico. Não reimplementa
compactação nem detecção de órfã.

## Regras da cascata

Ordem (importa quando dois usuários da lista se relacionam):

1. Resolver cada item da lista para um usuário da base de cadastro
   (CPF → nome → e-mail, como hoje).
2. Estrutura direta: para cada usuário resolvido, `find_traveler_structures`.
   Essas estruturas serão **excluídas** e saem do cálculo seguinte.
3. Compactação: nas **demais** estruturas, remover **todos** os CPFs da lista
   de `LoginAprovador_1..100` e do 2º nível, compactando e promovendo o 2º nível.
4. Órfã: estrutura (não excluída) que, depois do passo 3 com **todos** os CPFs
   da lista removidos, ficou sem nenhum aprovador. Estrutura excluída nunca
   conta como órfã.

Guarda: `delete_structures` leva todas as linhas do `AprovacaoId`; se algum id
a excluir tiver linha que não seja VIAJANTE de um CPF executável (outro
viajante, CCEMPRESA ou CPF em branco), a análise é recusada com
`ESTRUTURA_COMPARTILHADA` (400), sem exclusão parcial. Várias linhas do mesmo
viajante sob um id são permitidas.

Situações por usuário:

| Situação | Comportamento |
|---|---|
| Encontrado, ATIVO, com CPF | Entra na cascata |
| Encontrado sem CPF (vazio/ausente) | **Não executável.** Alerta exato: `Não foi possível mapear a Estrutura de Aprovação: Usuário encontrado no cadastro, mas não possui CPF registrado.` Corrigir o cadastro e refazer |
| Já INATIVO | Listado como `JA_INATIVO`, não executável (a ficha só considera ATIVO) |
| Não localizado | Listado como `NAO_LOCALIZADO` |
| Item que não é CPF (11 dígitos), e-mail nem nome completo | `NAO_LOCALIZADO` com alerta próprio; nunca entra na cascata |
| Nome com mais de um resultado | `candidatos` com CPF e e-mail; só entra na cascata depois de escolhido |
| CPF em duplicidade na lista | Uma vez só; reportado em `duplicados` |

Usuário sem CPF é bloqueado de propósito: inativá-lo sem limpar as estruturas
é exatamente o risco que esta funcionalidade elimina.

## Contrato das rotas

### `POST /api/inativacao/analisar` (multipart, sem efeito colateral)

Entrada: `cadastro` (arquivo), `estruturas` (arquivo), a lista como
`lista_text`, `lista` (arquivo) ou `itens` (JSON — a mesma leitura de hoje) e,
opcionalmente, `selecionados` (JSON com CPFs escolhidos entre os `candidatos`).

Saída `200`:

```jsonc
{
  "usuarios": [{
    "cpf": "12345678901", "cpfMascarado": "***.456.789-**",
    "nome": "João Silva", "email": "joao@empresa.com",
    "situacao": "EXECUTAVEL",   // | SEM_CPF | JA_INATIVO | NAO_LOCALIZADO | PENDENTE_SELECAO
    "alerta": null,             // texto exato acima quando SEM_CPF
    "estruturasViajante": ["APR001"],   // lista (um viajante pode ter mais de uma); [] se nenhuma
    "comoAprovador": [
      { "aprovacaoId": "APR010", "posicoes": [2], "acao": "COMPACTACAO" },
      { "aprovacaoId": "APR011", "posicoes": [1], "acao": "ORFA" }
    ],
    "candidatos": []            // preenchido em PENDENTE_SELECAO
  }],
  "resumo": { "executaveis": 1, "estruturasExcluidas": 1,
              "estruturasCompactadas": 1, "estruturasOrfas": 1,
              "duplicados": [] },
  "impressaoDigital": "<sha256>"
}
```

`comoAprovador[].acao` é por estrutura e por usuário; `ORFA` só aparece quando a
estrutura fica vazia depois de remover **todos** os CPFs da lista.

### `POST /api/inativacao/executar` (multipart)

Entrada: `cadastro`, `estruturas`, `cpfs` (JSON, já resolvidos — nunca texto
livre), `impressaoDigital`, `ignore_orphan_warning` (padrão `false`).

Comportamento: recalcula a análise a partir das planilhas; se a impressão
digital divergir, `409`; se houver órfãs e a confirmação não vier, `400`;
caso contrário devolve o ZIP e grava a auditoria. Se nenhum CPF for executável,
`400 NADA_A_EXECUTAR`. Nunca devolve sucesso sem gerar os arquivos.

### Impressão digital

SHA-256 do JSON canônico (chaves ordenadas, listas ordenadas) de: CPFs
executáveis, estruturas excluídas, estruturas compactadas (id + posições) e
estruturas órfãs. É semântica, não do arquivo: reenviar a mesma planilha salva
de outro jeito não invalida; trocar por outra que mude o impacto, sim.

### Saída do ZIP

- `saida_inativacao.xlsx`: ficha `DELETE` pelo motor atual, com a lista de CPFs.
- `estruturas_atualizadas.xlsx` (sempre presente; só o cabeçalho quando nenhuma
  estrutura é afetada): **todas** as linhas das estruturas afetadas,
  no mesmo formato do export de `/api/aprovacao/remover/export`;
  `Operacao=DELETE` nas excluídas, `UPDATE` nas compactadas. Estrutura que é
  excluída não recebe compactação.

## Erros

Cada erro tem `code` estável (o frontend decide pelo código, não pelo texto);
a mensagem é em português e não vaza detalhe interno.

| HTTP | code | Quando |
|---|---|---|
| 400 | `BASE_AUSENTE` / `ARQUIVO_INVALIDO` / `BASE_SEM_COLUNA` | falta arquivo, arquivo ilegível (extensão ou conteúdo), ou coluna ausente (CPF do viajante, `AprovacaoId`, `AprovacaoPor`, `LoginAprovador_*`, CPF no cadastro) |
| 400 | `LISTA_VAZIA` / `LISTA_GRANDE` | lista vazia ou acima de 500 itens |
| 400 | `NADA_A_EXECUTAR` | nenhum CPF executável |
| 400 | `ESTRUTURA_COMPARTILHADA` | um `AprovacaoId` a excluir reúne linhas de outros viajantes ou de outro tipo |
| 400 | `ORFAS_SEM_CONFIRMACAO` | há estruturas órfãs e `ignore_orphan_warning` não veio |
| 409 | `ANALISE_DIVERGENTE` | impressão digital diferente da recalculada |
| 413 | `ARQUIVO_GRANDE` | acima do teto das rotas |
| 500 | `ERRO_INTERNO` | mensagem genérica; detalhe só no log |

## Frontend

A aba vira um assistente de 3 etapas, no mesmo padrão do Cadastro
(`Stepper`, `FileDropzone`, `GenerationError`, `Modal`):

1. **Bases e lista** — duas áreas de upload (cadastro e estruturas) e a lista.
   Analisar só habilita com as duas bases.
2. **Impacto** — um cartão por usuário: quem será inativado, a estrutura direta
   a excluir e o impacto como aprovador (âmbar = compactação, vermelho =
   estrutura órfã). `SEM_CPF`, `JA_INATIVO` e `NAO_LOCALIZADO` aparecem com o
   motivo, fora da cascata. Homônimos são uma lista de escolha ("Aplicar
   seleção" refaz a análise) e **Continuar fica bloqueado enquanto houver nome
   repetido sem escolha**; para deixá-lo de fora, o operador o tira da lista.
3. **Confirmar** — confirmação consciente; com órfãs, uma segunda confirmação
   explícita. **Executar inativação** só habilita depois delas.

Os arquivos ficam em memória no navegador, então o `/executar` os reenvia sem o
operador anexar de novo (a tela mostra "arquivos já carregados"). Qualquer
mudança em arquivos ou lista descarta a análise e volta à etapa 1.

## Auditoria e privacidade

Eventos `inativacao_analise` e `inativacao_execucao` com: contagens, impressão
digital, situação por usuário e **CPF mascarado** (`***.456.789-**`). Nenhum nome,
e-mail nem CPF completo em `details`. Aplica-se a todo o histórico dessa aba.

## Escala e desempenho

- Uma passada sobre a base por conjunto de CPFs (não uma por CPF).
- Base normalizada e `melt` feitos uma vez e reaproveitados entre análise e
  execução; sem `iterrows` no código novo.
- Limites explícitos: lista de no máximo **500** itens; teto de upload próprio
  das duas rotas. Se a versão instalada do Flask não permitir
  `request.max_content_length` por rota, o teto global (`MAX_CONTENT_LENGTH`)
  sobe para 32 MB em `backend/core/config.py`. Erro sempre em português (`413`).
- Determinístico: as mesmas entradas geram a mesma saída; ZIP montado em
  memória; temporários apagados em `finally`.

## Testes

- **Negócio:** aprovador único (órfã); compactação; duas pessoas da lista que
  aprovam uma a outra; viajante que também aprova em outra estrutura; 2º nível
  com promoção; sem CPF; homônimos; já inativo; não localizado; CPF com
  pontuação e zero à esquerda; estrutura excluída nunca vira órfã.
- **Propriedades:** análise não altera as entradas (comparar antes/depois);
  execução repetida dá a mesma saída; impressão digital estável e sensível ao
  impacto.
- **Regressão do defeito antigo:** `/executar` lê as bases, gera os dois
  arquivos e registra auditoria; nunca responde sucesso sem arquivo.
- **Equivalência:** `remove_cpf_and_compact` (1 CPF) igual ao comportamento
  anterior; toda a suíte `test_aprovacao*.py` passa sem alteração.
- **Volume:** ~50 mil linhas × 100 usuários em até 15 s (teto folgado para o
  CI; se a medição real for muito menor, o teto é apertado na fatia 2).
- **API:** 400/409/413 com `code`; conteúdo do ZIP (abrir as duas planilhas e
  conferir `Operacao`); auditoria sem dado pessoal em claro.
- **Frontend:** hook e painel no Vitest; fluxo Playwright com caminho feliz e a
  trava de órfãs.

## Entrega em fatias

Cada fatia entra verde (backend `pytest`, frontend `vitest`, `tsc`, `ruff`,
`mypy`) antes da seguinte:

1. `ApprovalService`: passada única + novos métodos, com regressão.
2. `InactivationCascadeService`, máscara de CPF, impressão digital e `/analisar`.
3. `/executar` + ZIP + auditoria; remoção de `/process_inativacao`,
   `/preview_inativacao` e `/inativacao/buscar` com os testes que só existiam
   para elas (os de `InactivationService` ficam).
4. Frontend (assistente de 3 etapas) e remoção do hook antigo.
5. Documentação (`REGRAS_APROVACAO_INATIVACAO.md` deixa de dizer que as abas são
   independentes; `ARQUITETURA_MODERNIZADA.md`; README) e e2e.

## Fora de escopo (segunda fase, se fizer falta)

- Resolver a estrutura órfã no próprio painel (indicar novo aprovador,
  reaproveitando `replace_cpf`).
- Terceiro arquivo no ZIP com relatório de impacto por usuário.
- Alterar a base de cadastro do cliente (status `INATIVO`).

## Riscos

- **Refatorar `remove_cpf_and_compact`** mexe na aba Estruturas: mitigado pela
  suíte existente rodando sem alteração e pelo teste de equivalência.
- **Formato exato do export de estruturas** para a Argo: segue o export atual
  de remoção; confirmar com um arquivo real na fatia 3.
- **Coluna CPF na base de estruturas** pode faltar em bases antigas: vira
  `BASE_SEM_COLUNA` explícito, sem tentar adivinhar por nome.
