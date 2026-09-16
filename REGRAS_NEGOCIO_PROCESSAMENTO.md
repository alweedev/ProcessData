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
| Contratos de dados (`MODEL_COLS`, `FICHA_MAP`, `REQUIRED_OUTPUT_COLS`) | `backend/domain/rules.py` |
| Normalização de texto e CPF | `backend/shared/text_utils.py`, `backend/shared/cpf_utils.py` |

---

## 1. Fluxo de processamento

```
Arquivo (.xlsx / .xls / .xltx)
    │
    ├─ [1] Leitura + remoção de linhas de cabeçalho repetido
    ├─ [2] Mapeamento de colunas (FICHA_MAP, case/acento-insensitive)
    ├─ [3] Inicialização das 32 colunas do modelo (MODEL_COLS) + valores padrão
    ├─ [4] Nome/SobreNome a partir de NomeCompleto
    ├─ [5] Geração de Login (CPF ou Email) + fluxo SELF/FRONT
    ├─ [6] Sanitização de texto e normalização de booleanos
    ├─ [7] Validação por linha + validação geral (__geral__)
    ├─ [8] Desduplicação (Login + NomeCompleto)
    └─ [9] Remoção de linhas totalmente em branco → DataFrame final
```

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
| NOME | `Nome` |
| SOBRENOME (ATE 20 CARACTERES), SOBRENOME (limite 50 caracteres) | `SobreNome` |
| NOME COMPLETO, NomeCompleto, NOME COMPLETO (limite 50 caracteres) | `NomeCompleto` |
| EMAIL, E-MAIL, Email | `Email` |
| TELEFONE, Telefone | `Telefone` |
| EMPRESA (DO GRUPO), Empresa | `NomeEmpresa` |
| Centro de custo, Centro_de_Custo, CODIGO/CÓDIGO - CENTRO DE CUSTO | `CodigoCCustoEmpresa` |
| DESCRICAO - CENTRO DE CUSTO, Descrição Centro de Custo | `DescricaoCCustoEmpresa` |
| CARGO | `Cargo` |
| DEPARTAMENTO | `Departamento` |
| NIVEL, NÍVEL, NÍVEL (se aplicável) | `Nivel` |
| SOLICITANTE, SOLICITANTE? (S/N) | `Solicitante` |
| TERCEIRO, TERCEIRO? (S/N) | `Terceiro` |

Fonte única: `FICHA_MAP` em `backend/domain/rules.py`. Colunas da planilha que
não batem com nenhuma entrada do mapa são ignoradas.

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

### Colunas obrigatórias na saída (`REQUIRED_OUTPUT_COLS`)

`Login`, `NomeEmpresa`, `CodigoCCustoEmpresa`, `DescricaoCCustoEmpresa`,
`Email`, `NomeCompleto`, `Nome`, `SobreNome`, `CodigoIntegracao`,
`EmpresaCCustoParaUsuario`. Ausência de qualquer uma vira erro `__geral__`
(ver §5).

---

## 4. Nome, Login e fluxo (SELF / FRONT)

### Nome / SobreNome

Derivados de `NomeCompleto`: primeiro token → `Nome`, último token →
`SobreNome` (ambos sanitizados, máx. 20 caracteres). Não trata partículas
(da/de/dos) nem sobrenomes compostos — nome com 1 palavra só vira `Nome`,
`SobreNome` fica vazio.

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

Aplica-se a `Nome`, `SobreNome` (máx. 20), e sem limite a `NomeCompleto`,
`NomeEmpresa`, `DescricaoCCustoCliente`, `Cargo`, `Departamento`, `Cidade`,
`Estado`, `Endereco`:

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
qualquer outra coisa vira `"N"`.

**`Terceiro` é diferente**: se o valor contém dígitos, mantém só os dígitos
(é um ID de terceiro, não um flag); só cai no mapeamento S/N acima quando não
há dígito nenhum.

### `NroMatricula`

Mantém só os dígitos do valor de entrada.

---

## 6. Validação por linha (`ValidationService.validate_row`)

| # | Campo | Regra | Erro |
|---|---|---|---|
| 1 | `Solicitante` | precisa ser exatamente `S` ou `N` (maiúsculo) | `Solicitante obrigatório (deve ser S ou N)` |
| 2 | `CPF` | ver nota abaixo | `CPF deve ter 11 dígitos` |
| 3 | `Email` | se preenchido, precisa ter `@` e `.` depois do `@` | `Email inválido` |
| 4 | `NomeCompleto` | não pode ser vazio | `NomeCompleto vazio` |
| 5 | `Nivel` | ver autocorreção abaixo | `Nivel inválido, ajustado para vazio` |

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

- **Geral** (`ValidationService.validate_dataframe`): confere se todas as
  `REQUIRED_OUTPUT_COLS` existem no DataFrame e, quando há linhas, se
  `Login`, `Email`, `NomeCompleto`, `Nome`, `SobreNome`, `NomeEmpresa` não
  estão **inteiramente** vazias (coluna toda em branco = sinal de que a
  origem não trouxe aquele dado). Cada falha vira uma mensagem em
  `errors["__geral__"]`.
- **Desduplicação**: `drop_duplicates(subset=["Login", "NomeCompleto"], keep="first")`
  — mantém a primeira ocorrência.
- **Linhas em branco**: removidas ao final se `Login`, `NomeCompleto`, `CPF`
  e `Email` estiverem todos vazios.

---

## 8. Retorno do pipeline

```python
errors, df_final = ProcessingService.process_records_from_files(paths, login_choice, fluxo)
```

- Sucesso total: `errors == {}`, `df_final` com as 32 colunas do modelo.
- `errors[idx]`: mensagens da linha `idx`, concatenadas com `"; "`.
- `errors["__geral__"]`: falhas de colunas obrigatórias, concatenadas.
- `errors[caminho_do_arquivo]`: erro de leitura/I/O daquele arquivo.

---

## 9. Exemplos

### 9.1 Cadastro SELF — caminho feliz

Entrada: `CPF=123.456.789-00`, `NOME COMPLETO=João da Silva`,
`EMAIL=joao@empresa.com`, `EMPRESA=Empresa A`, `Centro de custo=CC001`,
`NIVEL=Operacional`, `SOLICITANTE=N`.

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
| `NomeCompleto` vazio | `NomeCompleto vazio` | preencher |
| CPF com mais de 11 dígitos, ou com menos de 10 dígitos reais | `CPF deve ter 11 dígitos` | conferir CPF (só 10 dígitos reais — 1 zero perdido pelo Excel — é aceito, ver §6) |
| Email sem `@`/`.` | `Email inválido` | formato `user@dominio.com` |
| `Solicitante` diferente de S/N | `Solicitante obrigatório (deve ser S ou N)` | usar S ou N |
| Coluna obrigatória ausente/vazia | `Coluna obrigatoria ausente: {col}` / `Coluna obrigatoria vazia: {col}` | adicionar/preencher a coluna |
| `Nivel` fora do esperado e sem substring reconhecível | `Nivel inválido, ajustado para vazio` | usar Operacional/Gerência/Diretoria |

---

## 11. Pontos de extensão

| Mudança | Onde mexer |
|---|---|
| Nova variação de nome de coluna | `FICHA_MAP` em `backend/domain/rules.py` |
| Nova validação por linha | `ValidationService.validate_row` |
| Novo fluxo além de SELF/FRONT | `ProcessingService.process_records_from_files`, bloco `fluxo_up` |
| Novo campo sanitizado como texto | lista de colunas em `ProcessingService.process_records_from_files` (loop de `sanitize_output_text`) |
