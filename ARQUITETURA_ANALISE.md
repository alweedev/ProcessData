# 🏗️ ANÁLISE ARQUITETURAL - Processamento de Planilhas

## 1. Estrutura de Módulos

```
backend/
├── processor.py          ← Transformação e normalização
├── validators.py         ← Validações de negócio
├── utils.py             ← Utilitários (CPF, texto, formatação)
└── api/
    ├── cadastro.py      ← Endpoint /api/process_cadastro
    ├── inativacao.py    ← Endpoints de inativação
    └── aprovacao.py     ← Endpoints de aprovação
```

---

## 2. Responsabilidades por Módulo

### 📄 processor.py

**Responsabilidades**:
1. Leitura de múltiplos formatos (DOCX, XLSX, XLS)
2. Mapeamento de colunas (normalização case-insensitive e sem acentos)
3. Extração de dados (extrair_docx, processamento Excel)
4. **Transformação de dados** (prioridade alta)
5. Desduplicação

**Funções Principais**:
- `extrair_docx(path)` → Extrai texto de Word
- `processar_registros_from_files(paths, login_choice, fluxo)` → Pipeline principal
- `drop_header_like_rows(df)` → Remove cabeçalhos repetidos
- `split_name_first_last(fullname)` → Divide nome/sobrenome
- `sanitize_output_text(v, maxlen)` → Normaliza texto
- `extract_digits_only(v)` → Extrai dígitos

**Dados Críticos**:
- `FICHA_MAP`: Mapa de colunas de entrada → campos padrão (23 variações mapeadas)
- `MODEL_COLS`: 33 colunas padrão de saída

### ✅ validators.py

**Responsabilidades**:
1. Validação de linha individual
2. Validação geral do DataFrame
3. Mapeamento inteligente de valores (Nivel)

**Funções Principais**:
- `validar_linha(reg)` → Valida 5 aspectos por linha
- `validar_dataframe_for_output(df)` → Valida colunas obrigatórias
- `validar_colunas_obrigatorias(df, required_cols)` → Helper genérico

**Validações Implementadas** (5):
1. Solicitante obrigatório (S/N)
2. CPF (se presente) com 11 dígitos
3. Email (se presente) com @ e .
4. NomeCompleto obrigatório
5. Nivel com mapeamento inteligente

### 🔧 utils.py

**Responsabilidades**:
1. Normalização de CPF
2. Normalização de texto
3. Formatação de saída
4. Validação de arquivo
5. Geração de nomes temporários

**Funções Principais**:
- `limpar_cpf_raw(cpf)` → Extrai 11 dígitos
- `format_cpf_for_output(digits)` → Formata XXX.XXX.XXX-XX
- `upper_no_accents(s)` → Maiúsculas sem acentos (Unicode NFKD)
- `validar_extensao_arquivo(filename)` → Valida .xlsx, .xls, .xltx
- `gerar_nome_arquivo_temporario(filename, folder)` → UUID + extensão

---

## 3. Fluxo de Processamento Detalhado

```
┌─────────────────────────────────────┐
│  Arquivo enviado (DOCX/XLS/XLSX)    │
└────────────┬────────────────────────┘
             │
             ├─→ [DOCX] extrair_docx() → texto → regex FICHA_MAP
             └─→ [XLS/XLSX] pd.read_excel() → drop_header_like_rows()
                                                ↓
                                    [Mapeamento de Colunas]
                                    (normalização case-insensitive)
                                                ↓
                    ┌─────────────────────────────────────┐
                    │  DataFrame com campos mapeados      │
                    └────────────┬────────────────────────┘
                                 │
                                 ├─→ Adiciona MODEL_COLS (se faltarem)
                                 ├─→ Atribui valores padrão
                                 │   - Operacao = "INSERT"
                                 │   - CodigoIntegracao = "AUT"
                                 │   - EmpresaCCustoParaUsuario = "S"
                                 ├─→ split_name_first_last() para Nome/SobreNome
                                 ├─→ Define Login (CPF ou EMAIL)
                                 │
                                 ├─→ [FLUXO SELF]
                                 │   └─ Todos os flags = "N"
                                 │
                                 ├─→ [FLUXO FRONT]
                                 │   ├─ ViajanteMasterNacional = "S"
                                 │   ├─ ViajanteMasterInternacional = "S"
                                 │   └─ Login prefixado com "FRONT"
                                 │
                                 ├─→ Sanitização de texto
                                 │   ├─ Nome/SobreNome (max 20)
                                 │   ├─ NomeCompleto
                                 │   ├─ Campos de descrição
                                 │   ├─ Email/Telefone → MAIÚSCULAS
                                 │
                                 ├─→ Normalização de numéricos
                                 │   ├─ NroMatricula → dígitos
                                 │   └─ Campos booleanos → S/N
                                 │
                    ┌────────────┴────────────────────────┐
                    │ [VALIDAÇÃO POR LINHA]               │
                    │ - Solicitante (obrigatório: S/N)    │
                    │ - CPF (11 dígitos se preenchido)    │
                    │ - Email (@ + .)                     │
                    │ - NomeCompleto (obrigatório)        │
                    │ - Nivel (mapeamento inteligente)    │
                    │                                      │
                    │ Retorna: errors[idx] = mensagens    │
                    └────────────┬────────────────────────┘
                                 │
                    ┌────────────┴────────────────────────┐
                    │ [VALIDAÇÃO GERAL]                   │
                    │ Verifica REQUIRED_OUTPUT_COLS       │
                    │ - Login, NomeEmpresa, CC            │
                    │ - Email, NomeCompleto, Nome         │
                    │ - SobreNome, CodigoIntegracao       │
                    │                                     │
                    │ Retorna: errors["__geral__"]        │
                    └────────────┬────────────────────────┘
                                 │
                    ┌────────────┴────────────────────────┐
                    │ [DESDUPLICAÇÃO]                     │
                    │ drop_duplicates(Login, NomeCompleto)│
                    └────────────┬────────────────────────┘
                                 │
                    ┌────────────┴────────────────────────┐
                    │ [NORMALIZAÇÃO FINAL]                │
                    │ - Preenchimento de campos bool      │
                    │ - Mapeamento S/N para todos os bool │
                    │ - Remoção de linhas em branco       │
                    │ - Conversão final de tipos          │
                    └────────────┬────────────────────────┘
                                 │
                                 ↓
                    ┌─────────────────────────────────┐
                    │ RETORNO                         │
                    │ (errors, df_final)              │
                    │                                 │
                    │ errors = {} ou {idx: msg}       │
                    │ df_final = DataFrame (33 cols)  │
                    └─────────────────────────────────┘
```

---

## 4. Matriz de Decisões

### Qual Campo Usar como Login?

| Escolha | Condição | Resultado | Validação |
|---------|----------|-----------|-----------|
| "CPF" (padrão) | `login_choice == "CPF"` | CPF formatado XXXXXXXXX-XX | 11 dígitos |
| "EMAIL" | `login_choice == "EMAIL"` | Email normalizado (MAIÚSCULAS) | @ + . |

### Qual Fluxo Usar?

| Fluxo | Caso de Uso | Flags de Acesso | Login Modificado |
|-------|-----------|-----------------|------------------|
| "SELF" (padrão) | Usuários internos simples | Todos = "N" | Não modificado |
| "FRONT" | Viajantes com acesso amplo | Master Nacional/Int = "S" | Prefixado "FRONT" |

### Quando um Valor é Inválido?

| Campo | Regra | Ação |
|-------|-------|------|
| Solicitante | ≠ S ou N | ✗ Erro |
| CPF | ≠ 11 dígitos | ✗ Erro |
| Email | Sem @ ou . | ✗ Erro |
| NomeCompleto | Vazio | ✗ Erro |
| Nivel | ∉ {OPER, GER, DIR} | ⚡ Tenta mapear, senão vazio |

---

## 5. Dependências de Dados

### Campos Críticos (Independentes)
```
Solicitante, NomeCompleto, Email
```

### Campos Derivados (Dependentes)
```
Nome, SobreNome
  ↑ Derivado de: NomeCompleto
  
Login
  ↑ Derivado de: CPF ou Email (depende de login_choice)
  ↑ Modificado por: Fluxo FRONT
  
Nivel
  ↑ Pode ser ajustado por: validador (mapeamento inteligente)
```

### Campos com Valores Padrão (Independentes)
```
Operacao = "INSERT"
CodigoIntegracao = "AUT"
EmpresaCCustoParaUsuario = "S"
Status = ""
Todos os flags booleanos = "N" (exceto em fluxo FRONT) - (exceto também não estejam preenchidas na planilha de entrada como SIM ou S, dessa forma integra como permitido a opção)
```

---

## 6. Pontos de Extensão

### 1. Adicionar Novo Mapeamento de Coluna
**Arquivo**: `processor.py` → `FICHA_MAP`

Exemplo:
```python
FICHA_MAP = {
    ...
    "NOVA_COLUNA": "NovoClaCampo",  # Adicionar aqui
    ...
}
```

### 2. Adicionar Nova Validação
**Arquivo**: `validators.py` → `validar_linha()`

Exemplo:
```python
def validar_linha(reg):
    msgs = []
    # ... validações existentes ...
    
    # Nova validação
    campo_novo = reg.get("CampoNovo", "").strip()
    if not campo_novo:
        msgs.append("CampoNovo obrigatório")
    
    return msgs
```

### 3. Adicionar Novo Fluxo
**Arquivo**: `processor.py` → seção de fluxos

Exemplo:
```python
if fluxo_up == "MEU_NOVO_FLUXO":
    df_final["MeuCampo"] = "X"
    # ... lógica específica ...
```

### 4. Adicionar Nova Transformação
**Arquivo**: `processor.py` → seção text_cols

Exemplo:
```python
text_cols = [
    # ... campos existentes ...
    "MeuCampoNovo",  # Adicionar aqui
]

for c in text_cols:
    if c in df_final.columns:
        if c == 'MeuCampoNovo':
            df_final[c] = df_final[c].apply(lambda v: minha_transformacao(v))
```

---

## 7. Questões Arquiteturais

### ❓ Por que tanta normalização?

1. **Dados sujos**: Usuários fornecem dados em formatos diferentes
2. **Integração Argo**: Plataforma downstream espera dados padronizados
3. **Consistência**: Evita erros causados por espaços, acentos, capitalização

### ❓ Por que MODEL_COLS é tão grande?

1. **Compatibilidade**: Precisa mapear todos os campos possíveis
2. **Flexibilidade**: Usuários podem fornecer dados parciais
3. **Saída consistente**: Sempre retorna mesma estrutura

### ❓ Por que desduplicação após validação?

1. **Performance**: Validar antes de desduplicar evita erro em linha já removida
2. **Lógica**: Mantém primeira ocorrência (válida)

### ❓ Por que Login pode ser modificado no fluxo FRONT?

1. **Integração**: Sistema Argo usa prefixo FRONT para identificar viajantes
2. **Rastreabilidade**: Facilita auditoria

---

## 8. Metriken de Qualidade

### Validações em Cascata

```
Total de Linhas
    │
    ├─→ Validação por linha (5 validações)
    │   └─→ Erros por linha
    │
    ├─→ Validação geral (colunas obrigatórias)
    │   └─→ Erros gerais
    │
    └─→ Desduplicação
        └─→ Linhas removidas
```

### Taxa de Sucesso

```
taxa_sucesso = (linhas_válidas / total_linhas) × 100%
```

---

## 9. Tratamento de Erros

### Estrutura de Retorno

```python
{
    # Erros por linha
    0: "Erro da linha 0",
    1: "Erro da linha 1",
    
    # Erro geral
    "__geral__": "Erro do DataFrame inteiro",
    
    # Erros de I/O
    "/caminho/arquivo.xlsx": "Erro ao ler arquivo"
}
```

### Exemplo Real

```python
{
    0: "Solicitante obrigatório (deve ser S ou N)",
    2: "NomeCompleto vazio; CPF deve ter 11 dígitos",
    "__geral__": "Coluna obrigatoria ausente: Email"
}
```

---

## 10. Padrões de Design Utilizados

| Padrão | Onde | Propósito |
|--------|------|----------|
| **Pipeline** | processor.py | Sequência de transformações |
| **Validator** | validators.py | Validações centralizadas |
| **Strategy** | login_choice, fluxo | Diferentes estratégias de processamento |
| **Factory** | FICHA_MAP | Mapeamento genérico de colunas |
| **Decorator** | sanitize_output_text | Normalização reutilizável |

---

## 11. Recomendações Arquiteturais

### ✅ O que está bem

1. **Separação de responsabilidades**: processor, validators, utils
2. **Mapeamento flexível**: FICHA_MAP permite variações
3. **Validações multicamadas**: linha + geral
4. **Tratamento de erros estruturado**: erros por índice

### ⚠️ Melhorias Possíveis

1. **Abstrair pipeline em classe**: Facilita testes e reutilização
2. **Extrair validações em objetos**: Facilita adicionar novas validações
3. **Usar enums para fluxos**: Em vez de strings
4. **Adicionar logging detalhado**: Para auditoria
5. **Schema validation**: Usar Pydantic para validação de tipos
6. **Testes parametrizados**: Cobrir todas as regras

---

## 12. Dependências Externas

| Biblioteca | Uso | Crítico |
|-----------|-----|---------|
| pandas | Manipulação de dados | ✅ Sim |
| openpyxl | Leitura de Excel | ✅ Sim |
| python-docx | Leitura de Word | ✅ Sim |
| unicodedata | Normalização Unicode | ✅ Sim |

---

## 📊 Resumo Estatístico

| Métrica | Valor |
|---------|-------|
| Colunas do Modelo | 33 |
| Colunas Obrigatórias | 10 |
| Validações por Linha | 5 |
| Variações de Coluna Mapeadas | 23+ |
| Fluxos Suportados | 2 (SELF, FRONT) |
| Formatos de Entrada | 3 (DOCX, XLSX, XLS) |
| Campos Booleanos | 8 |

