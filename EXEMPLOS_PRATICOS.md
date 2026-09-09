# 📚 EXEMPLOS PRÁTICOS - Processamento de Planilhas

> **Nota (set/2026):** os exemplos citam `processar_registros_from_files(...)`,
> que foi substituído por
> `ProcessingService.process_records_from_files(paths, login_choice, fluxo)`
> (mesma assinatura e mesmo resultado). Entrada aceita: `.xlsx`, `.xls`, `.xltx`.

## 1. Caso de Uso: Cadastro SELF (Padrão)

### Entrada (Excel)

```
| CPF | NOME COMPLETO | EMAIL | EMPRESA | CENTRO CUSTO | NIVEL |
|-----|---------------|-------|---------|--------------|-------|
| 123.456.789-00 | João da Silva | joao@empresa.com | Empresa A | CC001 | Operacional |
```

### Processamento

```python
errors, df_final = processar_registros_from_files(
    paths=["dados.xlsx"],
    login_choice="CPF",
    fluxo="SELF"
)
```

### Transformações Aplicadas

1. **Mapeamento de Colunas**:
   - NOME COMPLETO → NomeCompleto ✓
   - EMAIL → Email ✓
   - EMPRESA → NomeEmpresa ✓
   - CENTRO CUSTO → CodigoCCustoEmpresa ✓
   - NIVEL → Nivel ✓

2. **Geração de Campos**:
   - Nome = "JOAO" (primeiro token, max 20)
   - SobreNome = "SILVA" (último token, max 20)
   - Login = "123.456.789-00" (CPF formatado)

3. **Valores Padrão**:
   - Operacao = "INSERT"
   - CodigoIntegracao = "AUT"
   - EmpresaCCustoParaUsuario = "S"
   - Vip = "N"
   - ViajanteMasterNacional = "N"
   - ViajanteMasterInternacional = "N"
   - SolicitanteMaster = "N"
   - MasterAdiantamento = "N"
   - MasterReembolso = "N"

4. **Normalização de Texto**:
   - NomeCompleto: "JOAO DA SILVA" (maiúsculas, sem acentos)
   - NomeEmpresa: "EMPRESA A"
   - Nivel: "OPERACIONAL" (mapeado automaticamente)

5. **Validação por Linha**:
   - ✓ Solicitante: "N" (padrão preenchido, não está vazio)
   - ✓ CPF: "12345678900" (11 dígitos)
   - ✓ Email: válido (contém @ e .)
   - ✓ NomeCompleto: não vazio
   - ✓ Nivel: "OPERACIONAL" (válido)

### Saída (Linha Válida)

```python
{
    "Operacao": "INSERT",
    "UserId": "",
    "Login": "123456789-00",
    "CodigoCCustoCliente": "",
    "DescricaoCCustoCliente": "",
    "NomeEmpresa": "EMPRESA A",
    "CodigoCCustoEmpresa": "CC001",
    "DescricaoCCustoEmpresa": "",
    "EmpresaCCustoParaUsuario": "S",
    "NroMatricula": "",
    "Nome": "JOAO",
    "SobreNome": "SILVA",
    "NomeCompleto": "JOAO DA SILVA",
    "Email": "JOAO@EMPRESA.COM",
    "Telefone": "",
    "Cargo": "",
    "Departamento": "",
    "Nivel": "OPERACIONAL",
    "Endereco": "",
    "Cidade": "",
    "Estado": "",
    "CEP": "",
    "Solicitante": "N",
    "Terceiro": "N",
    "Vip": "N",
    "ViajanteMasterNacional": "N",
    "ViajanteMasterInternacional": "N",
    "SolicitanteMaster": "N",
    "MasterAdiantamento": "N",
    "MasterReembolso": "N",
    "CodigoIntegracao": "AUT",
    "Status": ""
}
```

---

## 2. Caso de Uso: Cadastro FRONT (Viajante)

### Entrada (Excel)

```
| EMAIL | NOME COMPLETO | EMPRESA |
|-------|---------------|---------|
| maria@empresa.com | Maria Santos Silva | Empresa B |
```

### Processamento

```python
errors, df_final = processar_registros_from_files(
    paths=["viajantes.xlsx"],
    login_choice="EMAIL",
    fluxo="FRONT"
)
```

### Transformações Específicas

1. **Login = EMAIL**:
   - Entrada: "maria@empresa.com"
   - Saída: "MARIA@EMPRESA.COM" (maiúsculas)
   - Com FRONT: "FRONTMARIA@EMPRESA.COM" ← **Prefixado**

2. **Fluxo FRONT**:
   - ViajanteMasterNacional = "S" ← **Diferença**
   - ViajanteMasterInternacional = "S" ← **Diferença**
   - Vip = "N"
   - SolicitanteMaster = "N"
   - MasterAdiantamento = "N"
   - MasterReembolso = "N"

3. **Divisão de Nome**:
   - NomeCompleto: "Maria Santos Silva"
   - Nome: "MARIA" (primeiro, max 20)
   - SobreNome: "SILVA" (último, max 20)

### Saída (Linha Válida FRONT)

```python
{
    "Operacao": "INSERT",
    "Login": "FRONTMARIA@EMPRESA.COM",  # ← Modificado
    "NomeCompleto": "MARIA SANTOS SILVA",
    "Nome": "MARIA",
    "SobreNome": "SILVA",
    "Email": "MARIA@EMPRESA.COM",
    "ViajanteMasterNacional": "S",  # ← FRONT
    "ViajanteMasterInternacional": "S",  # ← FRONT
    "NomeEmpresa": "EMPRESA B",
    ... (demais campos)
}
```

---

## 3. Caso de Uso: Dados Sujos com Autocorreção

### Entrada (Excel)

```
| CPF | Nome Completo | Email | Nível | Solicitante? |
|-----|---------------|-------|-------|--------------|
| 111222333-44 | João da Silva | joao@email.com | gerente | S |
```

### Processamento

Mesmo fluxo: `processar_registros_from_files()`

### Transformações com Autocorreção

1. **CPF**:
   - Entrada: "111222333-44" (sem zeros)
   - Limpeza: "11122233344" (10 dígitos)
   - ⚠️ Validação falha: "CPF deve ter 11 dígitos"

2. **Nível**:
   - Entrada: "gerente"
   - Normalização: "GERENTE"
   - Detecção: contém "GER"
   - ⚡ Autocorreção: "GERENCIA" ✓

3. **Solicitante**:
   - Entrada: "S"
   - Normalização: "S"
   - ✓ Válido (é S ou N)

### Saída com Erro

```python
errors = {
    0: "CPF deve ter 11 dígitos"
}

# Nota: A linha continua no DataFrame com Nivel corrigido para "GERENCIA"
# mas o processamento falha para essa linha devido ao CPF inválido
```

---

## 4. Caso de Uso: Desduplicação Automática

### Entrada (Excel)

```
| Login | Nome Completo |
|-------|---------------|
| 123.456.789-00 | João Silva |
| 123.456.789-00 | João Silva |  ← Duplicado
| 987.654.321-11 | Maria Santos |
```

### Processamento

```python
errors, df_final = processar_registros_from_files(
    paths=["duplicados.xlsx"],
    login_choice="CPF",
    fluxo="SELF"
)
```

### Lógica de Desduplicação

```python
df_final = df_final.drop_duplicates(
    subset=["Login", "NomeCompleto"],  # Critério
    keep="first"  # Mantém primeira ocorrência
)
```

### Saída

```
Total de linhas processadas: 3
Linhas após desduplicação: 2  ← 1 removida

| Login | Nome Completo |
|-------|---------------|
| 123456789-00 | João Silva |  ← Mantida
| 987654321-11 | Maria Santos |
```

---

## 5. Caso de Uso: Normalização Complexa

### Entrada (Dados Originais)

```
| CPF | NOME_COMPLETO | Email | Empresa | Descrição CC | Nível |
|-----|---------------|-------|---------|--------------|-------|
| 111.222.333-44 | João Pedrö dä Sïlva | joao@email.com | São José Inc. | COM AQUISICAO SFB (CO/N/NE) | oPeRaCiOnAl |
```

### Transformações Aplicadas

1. **CPF**:
   - Entrada: "111.222.333-44"
   - Limpeza: "11122233344"
   - Formatação: "111222333-44"

2. **Nome Completo**:
   - Entrada: "João Pedrö dä Sïlva"
   - Normalização: "JOAO PEDRO DA SILVA" (sem acentos)
   - Nome: "JOAO" (20 chars max)
   - SobreNome: "SILVA" (20 chars max)

3. **Email**:
   - Entrada: "joao@email.com"
   - Normalização: "JOAO@EMAIL.COM" (maiúsculas)

4. **Empresa**:
   - Entrada: "São José Inc."
   - Normalização: "SAO JOSE INC" (sem acentos)

5. **Descrição CC**:
   - Entrada: "COM AQUISICAO SFB (CO/N/NE)"
   - Normalização: "COM AQUISICAO SFB (CO/N/NE)" ← **Mantém estrutura**
   - (Remove acentos, mas preserva parênteses e barras)

6. **Nível**:
   - Entrada: "oPeRaCiOnAl"
   - Normalização: "OPERACIONAL"
   - Validação: ✓ Válido

### Saída Normalizada

```python
{
    "CPF": "111222333-44",
    "NomeCompleto": "JOAO PEDRO DA SILVA",
    "Nome": "JOAO",
    "SobreNome": "SILVA",
    "Email": "JOAO@EMAIL.COM",
    "NomeEmpresa": "SAO JOSE INC",
    "DescricaoCCustoEmpresa": "COM AQUISICAO SFB (CO/N/NE)",
    "Nivel": "OPERACIONAL",
    ...
}
```

---

## 6. Caso de Uso: Erros Múltiplos

### Entrada (Excel)

```
| CPF | Nome Completo | Email | Nivel |
|-----|---------------|-------|-------|
| 111 | | joao@email | invalido |
| | Maria | maria@email.com | GERENCIA |
```

### Processamento

Executa todas as 5 validações por linha.

### Saída com Múltiplos Erros

```python
errors = {
    0: "CPF deve ter 11 dígitos; Email inválido; NomeCompleto vazio; Nivel inválido, ajustado para vazio",
    1: "NomeCompleto vazio",  # Linha 2 (0-indexed = 1)
    "__geral__": ""  # Se colunas obrigatórias estiverem presentes
}
```

---

## 7. Caso de Uso: Mapeamento de Colunas Variadas

### Entrada (Excel com Nomes Diferentes)

```
| codigo_cpf | nome_pessoa | correio | agencia | centro_despesa | hierarquia |
|------------|-------------|---------|---------|-----------------|-----------|
| 111222333-44 | João Silva | joao@email.com | Empresa A | CC001 | Gerente |
```

### Mapeamento Aplicado

```python
# Normalização case-insensitive:
# "codigo_cpf" → "CODIGOCPF" → encontra "CPF" em FICHA_MAP
# "nome_pessoa" → "NOMEPESSOA" → encontra "NOMECOMPLETO"
# "correio" → "CORREIO" → encontra "EMAIL"
# "agencia" → "AGENCIA" → encontra "NOMEEMPRESA"
# "centro_despesa" → "CENTRODESPESA" → encontra "CODIGOCCUSTOEMPRESA"
# "hierarquia" → "HIERARQUIA" → encontra "NIVEL"

normalized_map = {
    "CODIGOCPF": "CPF",
    "NOMEPESSOA": "NomeCompleto",
    "CORREIO": "Email",
    "AGENCIA": "NomeEmpresa",
    "CENTRODESPESA": "CodigoCCustoEmpresa",
    "HIERARQUIA": "Nivel"
}
```

### Saída

Mesma estrutura padrão, colunas mapeadas corretamente.

---

## 8. Caso de Uso: Campos Opcionais vs Obrigatórios

### Cenário: Dados Mínimos

#### Entrada
```
| Email | Nome Completo | Centro Custo |
|-------|---------------|--------------|
| maria@email.com | Maria Santos | CC001 |
```

#### Processamento
- ❌ Sem CPF: NenhumErro (CPF é opcional se não fornecido)
- ❌ Sem empresa: Mas é obrigatória na saída!
- ❌ Sem Solicitante: Mas é obrigatória na saída!

#### Validação Geral Falha
```python
errors = {
    "__geral__": "Coluna obrigatoria ausente: Login; Coluna obrigatoria ausente: NomeEmpresa; Coluna obrigatoria ausente: CodigoIntegracao; Coluna obrigatoria ausente: EmpresaCCustoParaUsuario"
}
```

---

## 9. Caso de Uso: Importação Bem-Sucedida

### Entrada Completa e Válida

```
| CPF | Nome Completo | Email | Empresa | Centro Custo | Nivel | Solicitante |
|-----|---------------|-------|---------|--------------|-------|-------------|
| 123.456.789-00 | João Silva | joao@empresa.com | Empresa A | CC001 | Operacional | S |
| 987.654.321-11 | Maria Santos | maria@empresa.com | Empresa B | CC002 | Gerencia | N |
| 555.666.777-88 | Carlos Oliveira | carlos@empresa.com | Empresa A | CC001 | Diretoria | S |
```

### Processamento

```python
errors, df_final = processar_registros_from_files(
    paths=["vendas.xlsx"],
    login_choice="CPF",
    fluxo="SELF"
)
```

### Resultado

```python
errors = {}  # Vazio = sucesso total

len(df_final) = 3  # 3 linhas processadas

# Todas as validações passam:
# ✓ Todos os CPFs têm 11 dígitos
# ✓ Todos os emails são válidos
# ✓ Todos os nomes completos preenchidos
# ✓ Todos os níveis válidos (mapeados se necessário)
# ✓ Todos os solicitantes são S ou N
# ✓ Todas as colunas obrigatórias presentes
# ✓ Nenhuma duplicação Login + NomeCompleto
```

### Saída (Pronta para Argo)

```
3 linhas válidas
0 duplicatas removidas
0 erros
Status: PRONTO PARA CARGA
```

---

## 10. Mapeamento de Booleanos

### Entrada (Variações)

```
| Solicitante | Vip | Terceiro |
|-------------|-----|----------|
| Sim | Não | S |
| YES | NO | N |
| True | False | 1 |
| S | N | 0 |
| N | S | Sim |
```

### Transformação

```python
def map_bool_to_SN(v):
    s = upper_no_accents(str(v)).strip().upper()
    if s in ("S", "SIM", "YES", "Y", "TRUE", "1"):
        return "S"
    return "N"
```

### Saída

```
| Solicitante | Vip | Terceiro |
|-------------|-----|----------|
| S | N | S |
| S | N | N |
| S | N | S |
| S | N | N |
| N | S | S |
```

---

## 11. Estatísticas de Processamento

### Exemplo Real

```
Entrada:
- Arquivo: vendas_maio.xlsx
- Linhas: 150
- Colunas: 18

Processamento:
- Mapeamento: OK (12/18 colunas mapeadas)
- Transformação: OK
- Validação por linha: 5 erros encontrados
- Desduplicação: 3 duplicatas removidas
- Validação geral: OK

Saída:
- Linhas válidas: 142
- Linhas com erro: 5
- Duplicatas removidas: 3
- Taxa de sucesso: 94.7%

Status: ✓ PROCESSAMENTO CONCLUÍDO COM SUCESSO
```

