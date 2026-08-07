# 📋 Regras de Negócio - Processamento de Planilhas

## 📊 Visão Geral da Arquitetura

O backend processa planilhas através de um pipeline centralizado em **3 módulos principais**:

1. **`processor.py`** - Lógica de transformação de dados
2. **`validators.py`** - Validações de negócio
3. **`utils.py`** - Utilitários de normalização e formatação

---

## 🔄 FLUXO DE PROCESSAMENTO

```
Arquivo (DOCX/XLS/XLSX)
    ↓
[1] LEITURA & MAPEAMENTO DE COLUNAS
    ↓
[2] NORMALIZAÇÃO E TRANSFORMAÇÃO
    ↓
[3] VALIDAÇÃO POR LINHA
    ↓
[4] VALIDAÇÃO GERAL (DataFrame)
    ↓
[5] DESDUPLICAÇÃO
    ↓
[6] NORMALIZAÇÃO FINAL
    ↓
[7] SAÍDA VALIDADA
```

---

## 1️⃣ LEITURA E MAPEAMENTO DE COLUNAS

### Formatos Suportados
- **Excel**: `.xlsx`, `.xls`, `.xltx`
- **Word**: `.docx` (fichas individuais)
- ❌ Outros formatos são ignorados

### Mapas de Colunas (FICHA_MAP)

O sistema mapeia variações de nomes de colunas (case-insensitive, sem acentos) para campos padrão:

#### Campos de Identificação
| Entrada | Saída |
|---------|-------|
| CPF, CPF (SEM PONTOS) | CPF |
| MATRICULA, Matricula, NroMatricula | NroMatricula |

#### Dados Pessoais
| Entrada | Saída |
|---------|-------|
| NOME | Nome |
| SOBRENOME (ATE 20 CARACTERES) | SobreNome |
| NOME COMPLETO, NomeCompleto | NomeCompleto |
| EMAIL, E-MAIL | Email |
| TELEFONE | Telefone |

#### Dados Corporativos
| Entrada | Saída |
|---------|-------|
| EMPRESA (DO GRUPO), Empresa | NomeEmpresa |
| Centro de custo, Centro_de_Custo, CODIGO - CENTRO DE CUSTO | CodigoCCustoEmpresa |
| DESCRICAO - CENTRO DE CUSTO, Descrição Centro de Custo | DescricaoCCustoEmpresa |

#### Dados Funcionais
| Entrada | Saída |
|---------|-------|
| CARGO | Cargo |
| DEPARTAMENTO | Departamento |
| NIVEL, NÍVEL | Nivel |

#### Flags Booleanas
| Entrada | Saída |
|---------|-------|
| SOLICITANTE? (S/N) | Solicitante |
| TERCEIRO? (S/N), Terceiro | Terceiro |

### Remoção de Linhas Duplicadas (Cabeçalhos Repetidos)

Remove linhas onde >40% dos valores coincidem com o nome da coluna.

```python
def is_header(row):
    matches = 0
    for c in cols:
        val = str(row.get(c, "")).strip()
        if val.upper() == str(c).upper():
            matches += 1
    return (matches / len(cols)) > 0.4
```

---

## 2️⃣ NORMALIZAÇÃO E TRANSFORMAÇÃO

### 2.1 Inicialização de Colunas Padrão

**Colunas Padrão do Modelo (MODEL_COLS - 33 campos)**

```
Operacao, UserId, Login, CodigoCCustoCliente, DescricaoCCustoCliente,
NomeEmpresa, CodigoCCustoEmpresa, DescricaoCCustoEmpresa, EmpresaCCustoParaUsuario,
NroMatricula, Nome, SobreNome, NomeCompleto, Email, Telefone, Cargo, Departamento, Nivel,
Endereco, Cidade, Estado, CEP, Solicitante, Vip, ViajanteMasterNacional,
ViajanteMasterInternacional, SolicitanteMaster, MasterAdiantamento, MasterReembolso, Terceiro,
CodigoIntegracao, Status
```

**Valores Padrão Atribuídos**

| Campo | Valor Padrão | Descrição |
|-------|--------------|-----------|
| Operacao | "INSERT" | Operação padrão de cadastro |
| EmpresaCCustoParaUsuario | "S" | Associar empresa/CC ao usuário |
| CodigoIntegracao | "AUT" | Código de integração automático |
| Status | "" | Status vazio inicial |
| Solicitante | "N" | Padrão: não solicitante |
| Vip | "N" | Padrão: não VIP |
| ViajanteMasterNacional | "N" | Depende do fluxo |
| ViajanteMasterInternacional | "N" | Depende do fluxo |
| SolicitanteMaster | "N" | Padrão: não master |
| MasterAdiantamento | "N" | Padrão: não master |
| MasterReembolso | "N" | Padrão: não master |

### 2.2 Geração de Nome e Sobrenome

**Fonte**: `NomeCompleto` (campo obrigatório)

**Regra**: Divide em primeiro token (Nome) e último token (SobreNome)
- ❌ Não detecta partículas (da/de/dos)
- ❌ Não detecta sobrenomes compostos

```python
parts = [p for p in fullname.strip().split() if p]
if len(parts) == 1:
    first, last = parts[0], ""
else:
    first, last = parts[0], parts[-1]
```

**Limites de Caracteres**:
- Nome: máx 20 caracteres
- SobreNome: máx 20 caracteres

### 2.3 Geração de Login

**Escolhas Disponíveis** (parâmetro `login_choice`):

#### Opção 1: CPF (padrão)
```
Login = CPF formatado (XXX.XXX.XXX-XX)
Requer: campo CPF preenchido
Validação: 11 dígitos
```

#### Opção 2: EMAIL
```
Login = Email (normalizado para MAIÚSCULAS)
Requer: campo Email preenchido
```

### 2.4 Fluxos de Processamento

**Parâmetro**: `fluxo` (padrão: "SELF")

#### Fluxo SELF
Configura usuários com viagem doméstica:
```
Vip = "N"
ViajanteMasterNacional = "N"
ViajanteMasterInternacional = "N"
SolicitanteMaster = "N"
MasterAdiantamento = "N"
MasterReembolso = "N"
```

#### Fluxo FRONT
Configura viajantes com acesso completo:
```
ViajanteMasterNacional = "S"
ViajanteMasterInternacional = "S"
Vip = "N"
SolicitanteMaster = "N"
MasterAdiantamento = "N"
MasterReembolso = "N"

Login = "FRONT" + Login (sem espaços)
Exemplo: "João Silva" → "FONTjoaosilva" (com CPF) → "FONTxxx.xxx.xxx-xx"
```

### 2.5 Sanitização de Texto

**Aplica-se a campos textuais:**
```
Nome, SobreNome, NomeCompleto, NomeEmpresa,
DescricaoCCustoEmpresa, DescricaoCCustoCliente, Cargo,
Departamento, Cidade, Estado, Endereco
```

**Regras de Sanitização** (`sanitize_output_text()`):

1. Remove acentos (via `upper_no_accents()`)
2. Converte para MAIÚSCULAS
3. Remove caracteres especiais (mantém apenas: A-Z, 0-9, espaço, -, /, (), )
4. Remove espaços duplicados
5. Limita a caracteres (20 para Nome/SobreNome, ilimitado para outros)

**Exceção**: `DescricaoCCustoEmpresa`
- Apenas remove acentos
- Mantém estrutura original (parênteses, barras, etc.)
- Exemplo preservado: "COM AQUISICAO SFB (CO/N/NE)"

**Normalização de Email e Telefone**:
- Convertidos para MAIÚSCULAS
- Trim de espaços

### 2.6 Normalização de Campos Numéricos

**NroMatricula**:
- Extrai apenas dígitos
- Remove caracteres especiais

**Terceiro**:
- Se contiver dígitos: mantém dígitos
- Senão: mapeia Sim/Não → S/N

**Campos Booleanos** (`map_bool_to_SN()`):

Mapeia valores para S/N:
```
"S", "SIM", "YES", "Y", "TRUE", "1" → "S"
Tudo mais → "N"
```

Aplica-se a:
```
Solicitante, Terceiro, Vip, ViajanteMasterNacional,
ViajanteMasterInternacional, SolicitanteMaster,
MasterAdiantamento, MasterReembolso
```

---

## 3️⃣ VALIDAÇÃO POR LINHA

### Função: `validar_linha(reg)`

Validações executadas em cada registro:

#### 1. Solicitante (OBRIGATÓRIO)
```
✓ Deve ser 'S' ou 'N'
✗ Erro: "Solicitante obrigatório (deve ser S ou N)"
```

#### 2. CPF (CONDICIONAL)
```
✓ Se preenchido: deve ter 11 dígitos
✗ Erro: "CPF deve ter 11 dígitos"
⚠️ Warning: Se vazio mas esperado
```

Fonte do CPF: campo `CPF` ou `Login` (nessa ordem)

#### 3. Email (OPCIONAL COM VALIDAÇÃO)
```
✓ Se preenchido: deve conter @ e . após @
✗ Erro: "Email inválido"
⚠️ Warning: Se vazio mas esperado
```

Padrão: `email@dominio.com`

#### 4. NomeCompleto (OBRIGATÓRIO)
```
✓ Não pode estar vazio
✗ Erro: "NomeCompleto vazio"
```

#### 5. Nivel (CONDICIONAL COM AUTOCORREÇÃO)
```
✓ Valores aceitos: "OPERACIONAL", "GERENCIA", "DIRETORIA", ou vazio

⚡ Autocorreção (mapeamento inteligente):
- Contém "OPER" → ajusta para "OPERACIONAL"
- Contém "GER" → ajusta para "GERENCIA"
- Contém "DIR" → ajusta para "DIRETORIA"
- Outro texto → ajusta para vazio
  
✗ Erro: "Nivel inválido, ajustado para vazio"
```

---

## 4️⃣ VALIDAÇÃO GERAL (DataFrame)

### Função: `validar_dataframe_for_output(df)`

Verifica se o DataFrame contém todas as **COLUNAS OBRIGATÓRIAS DE SAÍDA**:

```
Login
NomeEmpresa
CodigoCCustoEmpresa
DescricaoCCustoEmpresa
Email
NomeCompleto
Nome
SobreNome
CodigoIntegracao
EmpresaCCustoParaUsuario
```

**Erro**: `"Coluna obrigatoria ausente: {coluna}"`

---

## 5️⃣ DESDUPLICAÇÃO

**Critério**: Combinação (Login, NomeCompleto)

```python
df_final = df_final.drop_duplicates(
    subset=["Login", "NomeCompleto"], 
    keep="first"
)
```

- Mantém primeira ocorrência
- Remove duplicatas exatas

---

## 6️⃣ NORMALIZAÇÃO FINAL

Aplicadas após validações:

### 6.1 Preenchimento de Campos Booleanos Faltantes

Se um campo booleano não existe na entrada → preenchido com "N"

### 6.2 Normalização de Valores Booleanos

Mapeia variações para S/N usando `map_bool_to_SN()`

### 6.3 Remoção de Linhas em Branco

Remove linhas onde todos os campos críticos estão vazios:
```
Campos críticos: Login, NomeCompleto, CPF, Email
(apenas se existirem)
```

### 6.4 Conversão de Tipos

- Objetos (strings) trimados e normalizados
- Email/Telefone convertidos para MAIÚSCULAS
- Login (EMAIL) convertido para MAIÚSCULAS
- DescricaoCCustoEmpresa normalizado (acentos removidos)
- NroMatricula com apenas dígitos

---

## 7️⃣ SAÍDA FINAL

### Estrutura do Retorno

```python
errors, df_final = processar_registros_from_files(paths, login_choice, fluxo)
```

#### Em Caso de Sucesso
- `errors`: dicionário vazio `{}`
- `df_final`: DataFrame com 33 colunas padrão, validado e normalizado

#### Em Caso de Erro
- `errors`: dicionário com índices de linha e mensagens
- `errors["__geral__"]`: erros gerais do DataFrame

#### Exemplo de Erro
```python
{
    0: "Solicitante obrigatório (deve ser S ou N); CPF deve ter 11 dígitos",
    2: "NomeCompleto vazio",
    "__geral__": "Coluna obrigatoria ausente: Email"
}
```

---

## 📐 MODELOS DE DADOS

### MODEL_COLS (33 Campos)

```python
[
    # Operação e Identificadores
    "Operacao",           # INSERT
    "UserId",             # ID do usuário (geralmente vazio)
    "Login",              # CPF formatado ou Email (obrigatório)
    "CodigoIntegracao",   # AUT (padrão)
    "Status",             # (vazio)
    
    # Identificação Pessoal
    "NroMatricula",       # Matrícula (opcional)
    "Nome",               # Primeiro nome (max 20 chars)
    "SobreNome",          # Último nome (max 20 chars)
    "NomeCompleto",       # Nome completo (obrigatório)
    "CPF",                # CPF com dígitos (11 dígitos)
    
    # Contato
    "Email",              # Email (obrigatório)
    "Telefone",           # Telefone (opcional)
    
    # Profissional
    "Cargo",              # Cargo (opcional)
    "Departamento",       # Departamento (opcional)
    "Nivel",              # OPERACIONAL, GERENCIA, DIRETORIA
    
    # Endereço
    "Endereco",           # Endereço (opcional)
    "Cidade",             # Cidade (opcional)
    "Estado",             # Estado (opcional)
    "CEP",                # CEP (opcional)
    
    # Centro de Custo
    "CodigoCCustoCliente",       # CC do cliente (opcional)
    "DescricaoCCustoCliente",    # Descrição do CC cliente
    "CodigoCCustoEmpresa",       # CC da empresa (obrigatório)
    "DescricaoCCustoEmpresa",    # Descrição CC empresa (obrigatório)
    "EmpresaCCustoParaUsuario",  # S/N (padrão S)
    
    # Dados da Empresa
    "NomeEmpresa",               # Nome da empresa (obrigatório)
    
    # Flags de Acesso (S/N)
    "Solicitante",               # É solicitante (obrigatório)
    "Terceiro",                  # É terceiro (S/N)
    "Vip",                       # É VIP (S/N)
    "ViajanteMasterNacional",    # Master nacional (S/N)
    "ViajanteMasterInternacional", # Master internacional (S/N)
    "SolicitanteMaster",         # Solicitante master (S/N)
    "MasterAdiantamento",        # Master de adiantamento (S/N)
    "MasterReembolso"            # Master de reembolso (S/N)
]
```

### REQUIRED_OUTPUT_COLS (10 Campos Obrigatórios)

```python
[
    "Login",
    "NomeEmpresa",
    "CodigoCCustoEmpresa",
    "DescricaoCCustoEmpresa",
    "Email",
    "NomeCompleto",
    "Nome",
    "SobreNome",
    "CodigoIntegracao",
    "EmpresaCCustoParaUsuario"
]
```

---

## 🔍 REGRAS ESPECIAIS

### Normalização de CPF

1. **Extração**: Remove tudo que não é dígito (`limpar_cpf_raw()`)
2. **Validação**: 11 dígitos obrigatórios
3. **Formatação**: XXX.XXX.XXX-XX (`format_cpf_for_output()`)

Exemplo:
```
Entrada: "123.456.789-00"
Extração: "12345678900"
Validação: ✓ 11 dígitos
Formatação: "123.456.789-00"
```

### Normalização de Texto (sem acentos)

Usa `upper_no_accents()`:
```
unicodedata.normalize("NFKD", s)
  .encode("ASCII", "ignore")
  .decode("utf-8")
```

Exemplo:
```
"São José" → "SAO JOSE"
"Açúcar" → "ACUCAR"
"Château" → "CHATEAU"
```

### Tratamento de Dados Vazios

- Strings vazias: `""` (vazio)
- None/NaN: convertidos para `""`
- Espaços em branco: trimados e normalizados

---

## ⚠️ CASOS DE ERRO COMUNS

| Situação | Erro | Resolução |
|----------|------|-----------|
| NomeCompleto vazio | "NomeCompleto vazio" | Preencher nome completo |
| CPF com <11 dígitos | "CPF deve ter 11 dígitos" | Verificar CPF |
| Email sem @ | "Email inválido" | Formato: user@domain.com |
| Solicitante não é S/N | "Solicitante obrigatório (deve ser S ou N)" | Usar S ou N |
| Coluna obrigatória ausente | "Coluna obrigatoria ausente: {col}" | Adicionar coluna |
| Nível inválido | "Nivel inválido, ajustado para vazio" | Usar OPERACIONAL, GERENCIA ou DIRETORIA |

---

## 📁 ESTRUTURA DE ARQUIVO DE ENTRADA

### Arquivo Excel Esperado

```
| CPF | NOME COMPLETO | EMAIL | CENTRO DE CUSTO | EMPRESA | ... |
|----|---|---|---|---|---|
| 123.456.789-00 | João Silva | joao@email.com | 001 - CC1 | Empresa A | ... |
| 987.654.321-11 | Maria Santos | maria@email.com | 002 - CC2 | Empresa B | ... |
```

### Arquivo Word (.docx) Esperado

```
CPF: 123.456.789-00
NOME COMPLETO: João Silva
EMAIL: joao@email.com
CENTRO DE CUSTO: 001 - CC1
EMPRESA (DO GRUPO): Empresa A
...
```

---

## 🎯 RESUMO ARQUITETURAL

| Aspecto | Descrição |
|---------|-----------|
| **Pipeline** | Sequencial com validações em camadas |
| **Entrada** | DOCX, XLSX, XLS |
| **Saída** | DataFrame normalizado + erros estruturados |
| **Validações** | 5 camadas (mapeamento, normalização, linha, geral, desdup) |
| **Colunas Modelo** | 33 campos padronizados |
| **Colunas Obrigatórias** | 10 campos mínimos |
| **Fluxos** | SELF (padrão) e FRONT |
| **Tratamento de Erros** | Por linha + geral, com warnings de campos vazios |

