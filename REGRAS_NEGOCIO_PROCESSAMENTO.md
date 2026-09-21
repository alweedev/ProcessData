# Regras de Negócio — Cadastro em Massa

Este documento descreve o pipeline de **cadastro** (`ProcessingService`): como uma
planilha de entrada vira a ficha padronizada pronta para carga na Argo. Para
**aprovação** e **inativação**, ver `REGRAS_APROVACAO_INATIVACAO.md`. Para a
visão de camadas/endpoints, ver `ARQUITETURA_MODERNIZADA.md`.

Código-fonte de referência:

| Responsabilidade | Módulo |
|---|---|
| Orquestração do pipeline | `backend/services/processing_service.py` (`ProcessingService.process_records_from_files`) |
| Validação por linha / geral | `backend/services/validation_service.py` (`ValidationService`) |
| Contratos de dados (`MODEL_COLS`, `FICHA_MAP`, `REQUIRED_FICHA_FIELDS`, `REQUIRED_OUTPUT_COLS`) | `backend/domain/rules.py` |
| Relatório da validação (aba "Validar" do Cadastro) | `backend/services/report_service.py`, `backend/api/analysis.py` |
| Normalização de texto e CPF | `backend/shared/text_utils.py`, `backend/shared/cpf_utils.py` |

---

## 1. Fluxo de processamento

```
Arquivo (.xlsx / .xls / .xltx)
    │
    ├─ [1] Leitura + remoção de linhas de cabeçalho repetido
    ├─ [2] Mapeamento de colunas (FICHA_MAP, case/acento-insensitive)
    ├─ [3] Inicialização das 32 colunas do modelo (MODEL_COLS) + valores padrão
    ├─ [4] Nome/SobreNome a partir de NomeCompleto (§4: nomes próprios | sobrenomes, sem cortar em 20)
    ├─ [5] Geração de Login (CPF ou Email) + fluxo SELF/FRONT
    ├─ [6] Sanitização de texto e normalização de booleanos
    ├─ [7] Validação por linha + validação geral (__geral__)
    ├─ [8] Desduplicação (Login + NomeCompleto), contando as removidas
    └─ [9] Remoção de linhas totalmente em branco → DataFrame final
```

A validação (passo 7) roda **antes** da normalização de sim/não (passo 6 só mexe nas
colunas de saída) e **não bloqueia** a geração: `/api/process_cadastro` gera o arquivo
com todas as linhas mesmo com erros; quem decide é o usuário (a tela pede confirmação
quando há linhas inválidas). `/api/analysis/summary` (a validação da tela) roda o
**mesmo pipeline** e devolve só o relatório (§8).

`.docx` nunca foi suportado (função inexistente); entrada aceita é só Excel.

---

## 2. Leitura e mapeamento de colunas

### Remoção de cabeçalho repetido

Se mais de 40% dos valores de uma linha coincidem (case-insensitive) com o
nome da própria coluna, a linha é descartada — cobre planilhas com o
cabeçalho colado de novo no meio dos dados.

### `FICHA_MAP` (variações de nome de coluna aceitas)

| Entrada (qualquer caixa/acento) | Campo padrão |
|---|---|
| CPF, CPF (SEM PONTOS) | `CPF` |
| MATRICULA, Matricula, MATRICULA (não obrigatório) | `NroMatricula` |
| NOME, NOME (limite 20 caracteres) | `Nome` |
| SOBRENOME (ATE 20 CARACTERES), SOBRENOME (limite 20 caracteres), SOBRENOME (limite 50 caracteres) | `SobreNome` |
| NOME COMPLETO, NomeCompleto, NOME COMPLETO (limite 50 caracteres), NOME COMPLETO (até 50 caracteres) | `NomeCompleto` |
| EMAIL, E-MAIL, Email | `Email` |
| E-MAIL (LOGIN), LOGIN (E-MAIL CORPORATIVO) | `EmailLogin` (prevalece sobre `Email`, ver §3) |
| TELEFONE, Telefone | `Telefone` |
| CELULAR - CONTATO, CELULAR CONTATO, CELULAR-CONTATO | `CelularContato` (2ª opção do telefone, §3) |
| Data de Nascimento, DATA NASCIMENTO, DATA DE NASC., DT NASCIMENTO, DT. NASCIMENTO, NASCIMENTO | `DataNascimento` |
| NÚMERO PASSAPORTE, NÚMERO PASSAPORTE (obrigatório para estrangeiro), Passport | `Passaporte` (dispensa o CPF, §3) |
| EMPRESA (DO GRUPO), Empresa | `NomeEmpresa` |
| Centro de custo, Centro_de_Custo, CODIGO/CÓDIGO - CENTRO DE CUSTO | `CodigoCCustoEmpresa` |
| DESCRICAO - CENTRO DE CUSTO, Descrição Centro de Custo | `DescricaoCCustoEmpresa` |
| CARGO | `Cargo` |
| DEPARTAMENTO | `Departamento` |
| NIVEL, NÍVEL, NÍVEL (se aplicável) | `Nivel` |
| SOLICITANTE, SOLICITANTE? (S/N) | `Solicitante` |
| TERCEIRO, TERCEIRO? (S/N) | `Terceiro` |

**Ficha em inglês** ("Registration form"): `Company` → `NomeEmpresa`, `Cost center code` →
`CodigoCCustoEmpresa`, `Cost center description` → `DescricaoCCustoEmpresa`, `First Name` →
`Nome`, `Last name (20 caracteres)` → `SobreNome`, `Full Name (50 caracteres)` → `NomeCompleto`,
`Mobile Number` → `Telefone`, `Position` → `Cargo`, `Department` → `Departamento`,
`Applicant? (Y/N)` → `Solicitante`, `Third Part? (Y/N)` → `Terceiro`, `Birth date` → `DataNascimento`.
São só apelidos para os mesmos campos: nenhuma regra nova.

Os nomes já existentes vêm de fichas antigas de clientes e **nunca são removidos**: novos apelidos só
se acrescentam. Fonte única: `FICHA_MAP` em `backend/domain/rules.py`. Colunas da planilha que
não batem com nenhuma entrada do mapa são ignoradas. Duas colunas para o mesmo campo: vale a última.

### Onde a planilha é lida

Só a **1ª aba** de cada arquivo, com o **cabeçalho na 1ª linha**. `.xlsx` é lido pelo `python-calamine`
(`backend/shared/excel_reader.py`), ~6x mais rápido que o `openpyxl`, com células idênticas (conferido nos
modelos de ficha dos clientes e em todos os tipos de célula); se falhar, cai no motor padrão do pandas.
`.xls` segue no `xlrd`.

Quando um arquivo não rende nenhum registro, o motivo é dito ao usuário: `A planilha está vazia.`,
`Nenhuma coluna da ficha foi reconhecida (colunas lidas: A, B, C). O cabeçalho precisa estar na 1ª
linha.` ou `A planilha não tem linhas de dados abaixo do cabeçalho.` — com `Só a 1ª aba ('X') é
lida.` no fim quando há mais abas.

---

## 3. Colunas do modelo e valores padrão

`MODEL_COLS` define as 32 colunas de saída, sempre nessa ordem. As que faltam
na entrada nascem vazias (`""`); em seguida alguns campos recebem valor fixo:

| Campo | Valor | Observação |
|---|---|---|
| `Operacao` | `"INSERT"` | sempre, no cadastro |
| `EmpresaCCustoParaUsuario` | `"S"` | fixo |
| `CodigoIntegracao` | `"AUT"` | fixo |
| `Status` | `""` | fica vazio |

### Campos obrigatórios da ficha (`REQUIRED_FICHA_FIELDS`)

São 8 (rótulo mostrado ao usuário → campo interno). **Em branco = linha inválida** (§6):

| Rótulo | Campo interno |
|---|---|
| CPF | `CPF` |
| Empresa | `NomeEmpresa` |
| Centro de custo - Código | `CodigoCCustoEmpresa` |
| Centro de custo - Descrição | `DescricaoCCustoEmpresa` |
| Nome completo | `NomeCompleto` |
| E-mail | `Email` |
| Telefone | `Telefone` |
| Data de nascimento | `DataNascimento` |

Todos os demais campos da ficha são **opcionais** (matrícula, cargo, departamento,
nível, sobrenome...). Os de sim/não (`Solicitante`, `Terceiro` etc.) nunca invalidam
a linha: ver §5.

`CPF` e `DataNascimento` são lidos **só para validar** (`SOURCE_ONLY_COLS`): não
fazem parte das 32 colunas de carga e não saem no arquivo gerado.

**E-mail de login.** Se a ficha traz `E-MAIL (LOGIN)` ou `LOGIN (E-MAIL CORPORATIVO)` (ARCOLOR,
IPMA, WEGHAUX trazem esta **e** a coluna de e-mail normal), o e-mail de login **prevalece**; o e-mail
normal só vale quando o de login está em branco na linha. Ficha só com a coluna de login não conta
como "coluna de e-mail ausente".

**CPF de estrangeiro (passaporte).** Se a linha tem número de passaporte, o CPF deixa de ser obrigatório
— **exceto no login por CPF**, em que o Login é gerado a partir dele (lá segue obrigatório). É por
linha: sem CPF e sem passaporte a linha continua inválida. Ficha sem coluna `CPF` mas com passaporte
não é "coluna ausente" (no login por e-mail), e todas as linhas serem de estrangeiros não gera
"coluna CPF vazia". O passaporte é só fonte: não vai para a carga.

**Telefone tem uma segunda opção.** Se o `Telefone` de uma linha está em branco e a
ficha traz o `CELULAR - CONTATO` (também `CELULAR CONTATO` / `CELULAR-CONTATO`), o
valor do contato é usado como `Telefone` **antes de validar** — então a linha não é
acusada e o valor sai no arquivo de carga. O `Telefone` preenchido sempre tem prioridade;
só continua em branco (linha inválida) quando nenhum dos dois foi preenchido. Se a
ficha nem tem a coluna `Telefone` mas tem o contato, isso não conta como coluna ausente.
O `CELULAR - CONTATO` em si não vai para a carga (também é `SOURCE_ONLY_COLS`).

`REQUIRED_OUTPUT_COLS` descreve a estrutura de saída (colunas que o modelo sempre
tem); não é usada para validar a ficha.

---

## 4. Nome, Login e fluxo (SELF / FRONT)

### Nome / SobreNome

As companhias aéreas exigem o nome **como no documento**. `Nome` leva **todos os nomes próprios** e
`SobreNome`, **todos os sobrenomes** (com partícula e sufixo), derivados de `NomeCompleto` — por exemplo
`Maria Clara da Silva Santos` → `Nome=MARIA CLARA`, `SobreNome=DA SILVA SANTOS`. As colunas `NOME` e
`SOBRENOME` da ficha só valem quando o nome completo está em branco. Implementação e testes:
`backend/shared/name_splitter.py`, `backend/tests/test_name_splitter.py`.

**Como separa** (tudo local e determinístico; nenhum nome sai da máquina):

1. **Regras.** Partícula (`DA DE DO DAS DOS DI DU DEL VAN VON DER LA E Y`...) faz parte do sobrenome que vem depois.
   Sufixo (`FILHO JUNIOR JR NETO SOBRINHO`...) gruda no sobrenome anterior (`DOS SANTOS FILHO`), mas não num nome
   próprio (`MARIA CLARA` + `NETO` deixa `NETO` como sobrenome). Nome composto com partícula (`MARIA DAS GRACAS`,
   `MARIA DE FATIMA`...) fica no `Nome`.
2. **Vocabulário.** Uma base embutida de nomes próprios e de sobrenomes, com frequência, diz de que lado da divisão
   cada palavra costuma ficar (`GABRIEL` é nome; como sobrenome é raro). O sistema escolhe a divisão de maior
   probabilidade.
3. **Confiança.** Cada divisão sai `alta` (segue sozinha) ou `baixa` (vai para a conferência): ambígua, uma só palavra,
   número no nome, sufixo sem sobrenome antes, partícula no início, nome com mais de 6 partes.

**Limite de 20 caracteres.** `Nome` e `SobreNome` **não são mais cortados** (cortar mudaria a identidade do passageiro
em silêncio). O que passar de 20 vai para a conferência com uma abreviação sugerida (nomes do meio viram inicial:
`FERNANDES DE ALBUQUERQUE CAVALCANTI` → `F DE A CAVALCANTI`); o `NomeCompleto` do documento nunca muda. A rota de
geração recusa (HTTP 400) enquanto restar nome acima do limite.

**Conferência de nomes** (etapa "Gerar" do assistente). A validação devolve `name_review` com só os nomes que precisam
de olhar. O usuário aceita ("Aceitar todas as sugestões") ou corrige; o "Gerar" só libera sem pendências. A decisão de
cada nome volta em `name_overrides` (`{"arquivo:linha": {nome_completo, nome, sobrenome, editado}}`) e só vale se o nome
completo da linha ainda for o mesmo — trocar a planilha descarta as decisões.

**Vocabulário aprendido.** Quando a geração dá certo, o que o usuário confirmou ensina o vocabulário: cada palavra do
`Nome` conta como nome próprio e cada palavra do `SobreNome`, como sobrenome. **Só palavras soltas com contagem** são
gravadas (nunca o nome completo nem o CPF), em `NAME_VOCAB_FILE` (padrão `~/.processdata/name_vocabulary.json`, fora do
repositório). Corrigir à mão pesa 3 e aceitar a sugestão pesa 1. Apagar o arquivo volta à base embutida. O histórico
(`/api/history`) só registra contagens, nunca nomes.

### Login (`login_choice`)

| Escolha | Resultado |
|---|---|
| `"CPF"` (padrão) | CPF formatado **`XXXXXXXXX-XX`** (9 dígitos + traço + 2 dígitos verificadores — **sem pontos**) |
| `"EMAIL"` | Email em MAIÚSCULAS |

> Formato real de `format_cpf_for_output()`: `123456789-00`, não
> `123.456.789-00`. Vale para cadastro **e** para os CPFs exibidos no fluxo de
> aprovação.

### Fluxo (`fluxo`, padrão `"SELF"`)

| Fluxo | `ViajanteMasterNacional`/`Internacional` | Login |
|---|---|---|
| `SELF` | `N` / `N` | inalterado |
| `FRONT` | `S` / `S` | prefixado `"FRONT"` (sem espaços) — ex.: `FRONT123456789-00` |

Em ambos os fluxos: `Vip`, `SolicitanteMaster`, `MasterAdiantamento`,
`MasterReembolso` = `"N"`.

---

## 5. Sanitização e normalização

### Texto (`sanitize_output_text`)

Aplica-se, **sem limite de tamanho**, a `Nome`, `SobreNome` (o que passa de 20 vai para a conferência de
nomes, §4, em vez de ser cortado), `NomeCompleto`, `NomeEmpresa`, `DescricaoCCustoCliente`, `Cargo`,
`Departamento`, `Cidade`, `Estado`, `Endereco`:

1. Remove acentos + maiúsculas (`upper_no_accents`, NFKD)
2. Mantém apenas `A-Z 0-9 espaço - / ( )`, remove o resto
3. Colapsa espaços duplicados

**Exceção `DescricaoCCustoEmpresa`**: só remove acentos, preserva vírgulas e
demais pontuação (ex.: `"COM AQUISICAO SFB (CO/N/NE), FASE 2"` continua igual
— usado como descrição livre pela Argo).

`Email`/`Telefone`: trim + MAIÚSCULAS (sem o filtro de caracteres).

### Campos booleanos → S/N

`Solicitante`, `Vip`, `ViajanteMasterNacional`, `ViajanteMasterInternacional`,
`SolicitanteMaster`, `MasterAdiantamento`, `MasterReembolso`: qualquer valor
em `{S, SIM, YES, Y, TRUE, 1}` (fold de acento antes de comparar) vira `"S"`;
qualquer outra coisa — `N`, `Não` **ou em branco** — vira `"N"`. Coluna ausente
da ficha também vira `"N"`. **Nada disso é erro de validação:** o campo é sempre
verificado e normalizado, nunca recusado.

**`Terceiro` é diferente**: se o valor contém dígitos, mantém só os dígitos
(é um ID de terceiro, não um flag); só cai no mapeamento S/N acima quando não
há dígito nenhum.

### `NroMatricula`

Mantém só os dígitos do valor de entrada.

---

## 6. Validação por linha (`ValidationService.validate_row`)

| # | Campo | Regra | Erro |
|---|---|---|---|
| 1 | Os 8 obrigatórios (§3) | não podem ficar em branco | `Campo obrigatório em branco: {rótulo}` (um por campo) |
| 2 | `CPF` | se preenchido, ver nota abaixo | `CPF deve ter 11 dígitos` |
| 3 | `Email` | se preenchido: **um endereço só**, sem espaço, `;` ou `,`; parte local não vazia; domínio com ponto, sem rótulo vazio e TLD de 2+ letras | `Email inválido` |
| 4 | `Nivel` | ver autocorreção abaixo | `Nivel inválido, ajustado para vazio` |

`Solicitante`, `Terceiro` e os demais campos de sim/não **não** são validados (§5). O
CPF é o da coluna `CPF` da ficha, mesmo quando o login é por e-mail.

> **Zero à esquerda vs. entrada incompleta.** `clean_cpf()` faz `zfill(11)`
> para restaurar o caso comum de Excel tratando CPF como número e derrubando
> **um** zero à esquerda (`"1234567890"`, 10 dígitos → `"01234567890"`).
> `validate_row` distingue isso de entrada realmente incompleta: só aceita o
> resultado do `zfill` se havia **pelo menos 10 dígitos** antes do
> preenchimento (`raw_cpf_digits()`); menos que isso gera
> `CPF deve ter 11 dígitos` mesmo depois de zero-preenchido (ex.: `"12345"`
> não vira um `"00000012345"` válido). Excesso de dígitos (12+) também
> falha — o `zfill` não trunca. Não há checagem de dígito verificador aqui —
> isso só existe no fluxo de aprovação (`ApprovalService.normalize_cpf_input`,
> módulo 11).

### Autocorreção de `Nivel`

Valores aceitos sem alteração: `OPERACIONAL`, `GERENCIA`, `DIRETORIA` (ou
vazio). Qualquer outro valor é normalizado (maiúsculas, sem acento) e, se
contiver `OPER`/`GER`/`DIR` como substring, é ajustado para o nível
correspondente; senão vira `""` e a linha registra o erro acima.

---

## 7. Validação geral e desduplicação

- **Geral** (`ValidationService.validate_dataframe`), sobre os 8 obrigatórios (§3):
  - coluna que a ficha **nem trazia** → `Coluna obrigatória ausente na ficha: {rótulo}`;
  - coluna presente mas **vazia em todas as linhas** → `Coluna obrigatória vazia na ficha: {rótulo}`
    (sinal de que a origem não trouxe o dado).

  Cada falha vira uma mensagem em `errors["__geral__"]`. Além disso, cada linha
  continua acusando o campo em branco (§6).
- **Desduplicação**: `drop_duplicates(subset=["Login", "NomeCompleto"], keep="first")`
  — mantém a primeira ocorrência. As removidas são **contadas antes** (só entre
  linhas com dados) e vão para o relatório como `duplicated_rows`.
- **Linhas em branco**: removidas se `Login`, `NomeCompleto`, `CPF` e `Email` estiverem
  todos vazios.
- **Erros só de quem sai no arquivo**: o erro de uma linha descartada (repetida ou em
  branco) é descartado junto, para `invalid_rows`/`valid_rows` baterem com o total.

---

## 8. Retorno do pipeline

```python
errors, df_final = ProcessingService.process_records_from_files(paths, login_choice, fluxo)
```

- Sucesso total: `errors == {}`, `df_final` com as 32 colunas do modelo.
- `errors[idx]`: mensagens da linha `idx` (índice interno do DataFrame), concatenadas com `"; "`.
- `errors["__geral__"]`: falhas de colunas obrigatórias, concatenadas.
- `errors[caminho_do_arquivo]`: por que aquele arquivo não pôde ser lido ou não rendeu registros (mesmo
  quando outros arquivos foram lidos: o arquivo não some do resultado sem aviso).
- `df_final.attrs` (metadados que não cabem nas colunas de carga; lidos pelo `ReportService`):
  `duplicated_rows`, `required_blank` (`{rótulo: nº de células em branco}`) e
  `row_labels` (`{idx: "Linha 4"}`, **só das linhas com erro**; com vários arquivos, `"Arquivo 2 · linha 4"`).
  Por que só das com erro: `attrs` é copiado em profundidade a cada operação do pandas, e um rótulo por
  linha (20 mil) chegou a custar 1/3 do tempo de exportar o arquivo.
- Vários arquivos de colunas diferentes: a célula que um arquivo não tem é **vazia** (não `NaN`), então é
  validada como em branco e não sai como o texto `NAN`.

### Parâmetros e erros das rotas

`login_choice` (`CPF`/`EMAIL`) e `fluxo` (`SELF`/`FRONT`): ausente ou vazio vale o padrão (`CPF`/`SELF`),
maiúscula/minúscula não importa, e valor desconhecido devolve **HTTP 400** (antes gerava um arquivo com Login
vazio). As mensagens de erro citam o **nome do arquivo enviado**, nunca o caminho temporário do servidor.
Planilha que não rende nenhuma linha devolve HTTP 400 com o motivo (§2) nas duas rotas
(`/api/analysis/summary` e `/api/process_cadastro`); arquivo quebrado enviado junto de arquivos bons entra
no `general_errors` do relatório.

### Desempenho

O pipeline não tem laço linha a linha: ele opera por coluna e calcula cada valor distinto uma vez (as fichas
repetem empresa, centro de custo, cargo...). Medido com 20.000 linhas: validar em ~2 s (antes ~12 s).
Gerar o arquivo ainda leva ~10 s nesse volume, quase tudo no `openpyxl` escrevendo as células.

### Relatório da validação (`POST /api/analysis/summary`)

`report` = `{total_rows, valid_rows, invalid_rows, duplicated_rows, general_errors, line_errors, line_details, required_blank}`:

- `total_rows`: linhas que **saem** no arquivo (já sem repetidas nem em branco);
  `valid_rows = total_rows - invalid_rows`.
- `line_errors`: `{"Linha 4": "msg; msg"}` — a **linha do Excel** (cabeçalho = linha 1;
  a ordem dos arquivos é a da lista enviada).
- `line_details`: `[{"label": "Linha 4", "nome": "ANA SOUZA", "erros": ["msg", "msg"], "sem_preenchimento": false}]` — o
  mesmo, uma entrada por linha com problema, **com o passageiro** e cada problema separado. `sem_preenchimento` é
  `true` quando a linha não tem **nenhum** dos obrigatórios (só o nome): serve só para a tela dizer isso numa frase; a
  linha continua com pendência e a regra de validação não muda.
- `required_blank`: um contador por campo obrigatório, com o nome que aparece na ficha.

**Como a tela mostra** (última etapa do Cadastro, sempre **acima** dos botões; a validação roda sozinha ao chegar em
"Gerar" e refaz quando fichas, login ou fluxo mudam — por isso não há botão de revalidar):

- **Tudo certo:** uma linha só ("Tudo certo: 5 cadastros prontos para gerar"); os números (Linhas, Válidas, Inválidas,
  Duplicadas) ficam atrás de "Ver detalhes".
- **Só aviso** (linhas repetidas removidas): a mesma linha, dizendo o que foi feito.
- **Com pendência:** **um cartão só** — título ("4 de 5 cadastros têm pendência. 1 pronto"), a orientação ("Corrija a
  ficha ou gere assim mesmo (pediremos sua confirmação)") e, dentro dele, o detalhe, sem repetir a contagem:
  - problema da ficha inteira (coluna ausente/vazia) em uma frase — o "X em branco" que ele já explica não é repetido por
    passageiro;
  - o que falta em *todas* as linhas dito uma vez ("Em todos os 3: Telefone em branco, …") e cada passageiro uma vez
    ("Linha 5 · NOME — também: E-mail inválido"); com problemas diferentes por linha, cada uma lista os seus;
  - linhas sem nenhum obrigatório num bloco à parte ("6 linhas sem nenhum campo obrigatório — parece que não foram
    preenchidas"), fora da conta do que é comum;
  - cada lista mostra 5 e abre o resto sob demanda. Para trocar a ficha há o botão **"Trocar arquivos"** da linha Fichas (no
    resumo acima): abre a escolha de arquivos ali mesmo, sem voltar à 1ª etapa; login e fluxo ficam como estão e a
    validação roda de novo sozinha. (Os botões "Alterar" de login e fluxo levam à etapa correspondente.)
  O "Gerar cadastro" fica em âmbar (destaque sem fingir que está tudo certo) e pede confirmação.
- **Planilha recusada** (ilegível, sem colunas conhecidas, vazia): um cartão só, no lugar do veredito ("Não foi possível
  validar a planilha"), com o motivo por arquivo (o servidor manda `errors: {arquivo: motivo}` junto do `error`, na
  validação e na geração), o que fazer e, recolhida, a mensagem original em "Detalhes técnicos". Como gerar falharia pelo
  mesmo motivo, o "Gerar cadastro" fica desabilitado (com o motivo no tooltip) e não aparece um segundo cartão de erro.
- **Falha ao gerar** (rede, erro do servidor, ...): o mesmo cartão, abaixo dos botões ("Não foi possível gerar o
  cadastro"), com texto claro e o que fazer.

`invalid_rows` e `general_errors` pedem confirmação para gerar; `duplicated_rows` (já removidas do arquivo) só informa.

---

## 9. Exemplos

### 9.1 Cadastro SELF — caminho feliz

Entrada: `CPF=123.456.789-00`, `NOME COMPLETO=João da Silva`,
`EMAIL=joao@empresa.com`, `EMPRESA=Empresa A`, `Centro de custo=CC001`,
`Descrição Centro de Custo=ADMINISTRATIVO`, `TELEFONE=11999990000`,
`Data de Nascimento=12/05/1990`, `NIVEL=Operacional`, `SOLICITANTE=N`
(os 8 obrigatórios preenchidos).

> Reparo: `Centro de custo` (com "de") é a grafia reconhecida por
> `FICHA_MAP` — uma variação como `CENTRO CUSTO` (sem "de") não bate com
> nenhuma chave e fica de fora do mapeamento (ver §9.4).

Saída (linha válida, `login_choice="CPF"`, `fluxo="SELF"`):

```python
{
    "Operacao": "INSERT", "Login": "123456789-00",
    "NomeEmpresa": "EMPRESA A", "CodigoCCustoEmpresa": "CC001",
    "Nome": "JOAO", "SobreNome": "SILVA", "NomeCompleto": "JOAO DA SILVA",
    "Email": "JOAO@EMPRESA.COM", "Nivel": "OPERACIONAL",
    "EmpresaCCustoParaUsuario": "S", "CodigoIntegracao": "AUT",
    "Vip": "N", "ViajanteMasterNacional": "N", "ViajanteMasterInternacional": "N",
    # ... demais campos do modelo, vazios ou "N"
}
```

### 9.2 Cadastro FRONT — viajante

Entrada: `EMAIL=maria@empresa.com`, `NOME COMPLETO=Maria Santos Silva`,
`login_choice="EMAIL"`, `fluxo="FRONT"`.

Diferenças-chave: `Login = "FRONTMARIA@EMPRESA.COM"` (prefixado, sem
espaços), `ViajanteMasterNacional = "S"`, `ViajanteMasterInternacional = "S"`.

### 9.3 Dados sujos com autocorreção e erro

Entrada: `CPF=123456789012` (12 dígitos — número digitado errado),
`Nível=gerente`, `SOLICITANTE? (S/N)=S`.

- `Nivel`: `"gerente"` → contém `GER` → autocorrigido para `"GERENCIA"`
  (a correção é aplicada mesmo assim, a linha só falha pelo CPF).
- `CPF`: 12 dígitos → excede 11 → `errors[idx] = "CPF deve ter 11 dígitos"`.
  Um CPF **bem curto** (ex.: `"12345"`, 5 dígitos reais) cai no mesmo erro —
  só 10 dígitos reais (1 zero perdido pelo Excel) são aceitos (ver §6).
- `Solicitante`: `"S"` → válido.

### 9.4 O que a normalização de coluna faz (e o que não faz)

A normalização em `ProcessingService` é só `upper_no_accents(col).strip()` —
maiúsculas + remoção de acento + trim nas pontas. Ela **não** remove espaços
internos, `_` ou `-`, e não faz matching semântico:

| Coluna de entrada | Normalizado | Casa com `FICHA_MAP`? |
|---|---|---|
| `nome completo` / `Nome Completo` | `NOME COMPLETO` | ✓ → `NomeCompleto` |
| `NOME_COMPLETO` (underscore) | `NOME_COMPLETO` | ✗ — chave no mapa é com espaço, não `_` |
| `nome_pessoa` | `NOME_PESSOA` | ✗ — variação não cadastrada |
| `e-mail` | `E-MAIL` | ✓ → `Email` |

Ou seja: `FICHA_MAP` é uma tabela fixa de sinônimos exatos (após o fold de
caixa/acento), não um matching aproximado. Para aceitar uma variação nova de
nome de coluna (ex.: `NOME_COMPLETO` com underscore), ela precisa ser
adicionada como chave literal em `FICHA_MAP` (§2).

---

## 10. Casos de erro comuns

| Situação | Erro | Resolução |
|---|---|---|
| Qualquer dos 8 obrigatórios em branco (CPF, empresa, centro de custo código/descrição, nome completo, e-mail, telefone — depois de tentar o `CELULAR - CONTATO` —, data de nascimento) | `Campo obrigatório em branco: {rótulo}` | preencher |
| CPF com mais de 11 dígitos, ou com menos de 10 dígitos reais | `CPF deve ter 11 dígitos` | conferir CPF (só 10 dígitos reais — 1 zero perdido pelo Excel — é aceito, ver §6) |
| Email sem `@`/`.` | `Email inválido` | formato `user@dominio.com` |
| `Solicitante`/`Terceiro` com `Sim`, `Não` ou em branco | (sem erro) | vira `S` / `N` / `N` — nada a corrigir |
| Coluna obrigatória ausente/vazia | `Coluna obrigatória ausente na ficha: {rótulo}` / `Coluna obrigatória vazia na ficha: {rótulo}` | adicionar/preencher a coluna |
| `Nivel` fora do esperado e sem substring reconhecível | `Nivel inválido, ajustado para vazio` | usar Operacional/Gerência/Diretoria |

---

## 11. Pontos de extensão

| Mudança | Onde mexer |
|---|---|
| Nova variação de nome de coluna | `FICHA_MAP` em `backend/domain/rules.py` |
| Novo campo obrigatório | `REQUIRED_FICHA_FIELDS` em `backend/domain/rules.py` (e, se a ficha o traz mas ele não vai para a carga, `SOURCE_ONLY_COLS`) |
| Nova validação por linha | `ValidationService.validate_row` |
| Novo fluxo além de SELF/FRONT | `ProcessingService.process_records_from_files`, bloco `fluxo_up` |
| Novo campo sanitizado como texto | lista de colunas em `ProcessingService.process_records_from_files` (loop de `sanitize_output_text`) |
