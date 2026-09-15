# 🔐 REGRAS DE NEGÓCIO - Aprovação e Inativação

## 1. MÓDULO APROVAÇÃO (aprovacao.py)

### 1.1 Objetivo

Gerenciar estruturas de aprovação em cadeia, permitindo:
- Substituir aprovadores em estruturas (fluxos de viagem/reembolso)
- Remover aprovadores
- Identificar estruturas que ficarão sem aprovador
- Gerar relatórios de impacto

---

### 1.2 Conceitos-Chave

#### Estrutura de Aprovação
Uma estrutura representa um **fluxo de aprovação** na plataforma Argo:
- Identificada por `AprovacaoId` (chave única)
- Pode ter múltiplos aprovadores (`LoginAprovador_1` até `LoginAprovador_100`)
- Pode ser aprovada por **VIAJANTE** ou **CCEMPRESA** (Centro de Custo)

#### Níveis de Aprovação
```
Nível 1: Aprovador principal
Nível 2: Aprovador secundário (pode estar ausente)
```

#### Aprovador por
- **VIAJANTE**: Aprovação designada a uma pessoa específica
- **CCEMPRESA**: Aprovação designada a um centro de custo (departamento)

---

### 1.3 Regras de Validação

#### 1. CPF do Aprovador (OBRIGATÓRIO) — enforçado

```
✓ Informado
✓ 11 dígitos
✓ Dígito verificador válido (módulo 11)
✓ Existe na base de usuários
✓ O usuário está com Status = ATIVO (quando a base tem coluna Status)

✗ "Informe um CPF para o aprovador."
✗ "CPF inválido. Informe 11 dígitos."
✗ "CPF inválido (dígito verificador)."
✗ "CPF não encontrado na base de usuários."
✗ "Aprovador não está ATIVO na base de usuários (Status: '...')."
```

> A validação de dígito verificador é aplicada **apenas** no fluxo de aprovação;
> cadastro e inativação seguem com checagem de comprimento.

#### 2. Base de Usuários (OBRIGATÓRIO) — enforçado

```
✓ Coluna 'CPF' (detecção case-insensitive, sem acento/_/-)
✓ Coluna de nome: 'NomeCompleto' OU 'Nome' (+ opcional 'SobreNome')

✗ "Base de usuários não contém coluna 'CPF'."
✗ "Base de usuários não contém coluna de nome ('NomeCompleto' ou 'Nome')."
✗ "Falha ao ler base de usuários: {erro}"
```

#### 3. Detecção de Colunas (AUTOMÁTICO)

O sistema detecta automaticamente as colunas:
```
Obrigatórios:
- AprovacaoId: chave da estrutura
- LoginAprovador_1..100: coluna com login do aprovador

Opcionais:
- AprovacaoPor: tipo de aprovação (VIAJANTE ou CCEMPRESA)
- Aprovacao: nome da aprovação
- Tipo: tipo de fluxo
- Valor: valor da estrutura
- DescricaoCCusto: descrição do CC
- CodigoCCusto: código do CC
- NomeViajante: nome de quem viaja
```

#### 4. Validação de Estrutura Sem Aprovador

Verifica estruturas que **ficarão vazias** após remoção do aprovador:

```
Cenário: Se removerem um aprovador e a estrutura não tiver fallback
Resultado: Marca como "estrutura_sem_aprovador: true"
Ação: Deve notificar usuário antes de salvar
```

---

### 1.4 Fluxo de Aprovação

#### Substituição de Aprovador

> Endpoints: `POST /api/aprovacao/substituir/preview` e `POST /api/aprovacao/substituir/export`
> (`backend/api/aprovacao.py`, lógica em `ApprovalService.replace_cpf` / `check_new_approver_duplicates`).

Cenário real de negócio: um aprovador sai e outro assume o lugar dele nas
mesmas estruturas — não existe um fluxo de "inserir um aprovador do zero em
estruturas arbitrárias" (isso exigiria decidir em quais estruturas inserir e
em qual posição, o que não tem uma regra de negócio definida; ver PDF de
carga da Argo para o formato completo da planilha, caso esse fluxo venha a
ser necessário no futuro).

```
[1] Usuário envia:
    - CPF do aprovador atual (que está saindo)
    - CPF do novo aprovador (que está entrando)
    - Arquivo base com estruturas (base_file)
    - Arquivo de usuários (users_file)

[2] Validação:
    - Os dois CPFs têm 11 dígitos e dígito verificador válido
    - Os dois CPFs existem na base de usuários e estão ATIVO
    - CPF novo != CPF atual
    - Base tem coluna AprovacaoId e colunas LoginAprovador_1..100

[3] Detecção:
    - Localiza todas as estruturas onde o CPF atual aparece
      (LoginAprovador_1..100 e, opcionalmente, LoginAprovador_SEGUNDO_NIVEL)
    - Verifica se o novo CPF já é aprovador em alguma dessas estruturas
      (duplicidade) — reportado em `estruturasComDuplicidade`

[4] Preview (antes de executar):
    - Mostra estruturas afetadas, posições onde o CPF atual aparece
    - Mostra estruturas que já têm o novo CPF (`teraDuplicidade`)

[5] Confirmação do usuário:
    - Revisa impacto; se houver duplicidade, decide se continua
      (`ignore_duplicate_warning=true`) ou cancela

[6] Execução:
    - Substitui o CPF atual pelo novo em LoginAprovador_1..100, na MESMA
      posição (sem compactar) — ao contrário da remoção, a estrutura não
      muda de forma, só o aprovador muda
    - Se `replace_second_level=true` e o CPF atual está no
      LoginAprovador_SEGUNDO_NIVEL, substitui lá também (padrão: false,
      mantém o segundo nível como está)
    - Gate: se alguma estrutura selecionada já tiver o novo CPF como
      aprovador (duplicidade), retorna aviso (HTTP 400 + warning) e só
      prossegue com ignore_duplicate_warning=true

[7] Export:
    - Exporta TODAS as linhas das estruturas alvo (a Argo precisa da
      estrutura completa)
    - Operacao = "UPDATE" apenas nas linhas efetivamente alteradas
```

#### Remoção de Aprovador

```
[1] Preview (validação antes de executar):
    - Mostra estruturas que serão afetadas
    - Mostra posições onde aparece
    - Identifica estruturas que ficarão sem aprovador
    - Agrupa por tipo de aprovação

[2] Confirmação do usuário:
    - Revisa impacto
    - Confirma remoção

[3] Execução:
    - Remove o CPF de LoginAprovador_1..100 e compacta à esquerda (sem vazios no meio)
    - Se o 1º nível ficou vazio e há LoginAprovador_SEGUNDO_NIVEL preenchido
      (que não seja o próprio CPF removido): promove o 2º nível para o 1º e
      esvazia o 2º nível (§2.4 Fase 3)
    - Se remove_second_level=true e o 2º nível é o CPF removido: apaga o 2º nível
    - Gate: se alguma estrutura selecionada ficar sem NENHUM aprovador, retorna
      aviso (HTTP 400 + warning) e só prossegue com ignore_empty_warning=true

[4] Export:
    - Exporta TODAS as linhas das estruturas alvo (a Argo precisa da estrutura
      completa)
    - Operacao = "UPDATE" apenas nas linhas efetivamente alteradas;
      as demais linhas ficam sem carimbo
```

---

### 1.5 Estrutura de Dados Retornada

#### Preview Response

```python
{
    "aprovador_nome": "João Silva",
    "aprovador_cpf": "123.456.789-00",
    "total_estruturas": 25,
    "total_ocorrencias": 42,
    "estruturas_por_tipo": {"VIAJANTE": 15, "CCEMPRESA": 10},
    "estruturas_sem_aprovador": [{"aprovacao_id": "APR001", "tipo": "VIAJANTE", "motivo": "Será única aprovadora"}],
    "estruturas": [
        {
            "aprovacao_id": "APR001",
            "aprovacao_por": "VIAJANTE",
            "aprovacao": "Aprovação 1",
            "tipo": "Reembolso",
            "valor": "1000.00",
            "viajante_nome": "Maria",
            "cc_codigo": "CC001",
            "cc_descricao": "Centro A",
            "posicoes": [1, 5, 8],
            "segundo_nivel": true,
            "ficara_sem_aprovador": false,
        }
    ],
}
```

---

### 1.6 Casos de Erro

| Situação | Erro | Resolução |
|----------|------|-----------|
| CPF vazio | "Informe um CPF para o aprovador" | Preencher CPF |
| CPF inválido | "CPF inválido. Informe 11 dígitos" | Formato: XXX.XXX.XXX-XX |
| CPF não existe | "CPF não encontrado na base de usuários" | Verificar se usuário existe |
| Base sem CPF | "Base não contém coluna 'CPF'" | Adicionar coluna CPF |
| Base vazia | "Base de usuários vazia" | Fornecer base válida |

---

## 2. MÓDULO INATIVAÇÃO (inativacao.py)

### 2.1 Objetivo

Buscar e inativar usuários em massa:
- Buscar por CPF, Email ou Nome Completo
- Validar antes de inativar
- Gerar relatórios de inativação
- Remover estruturas de aprovação associadas

---

### 2.2 Critérios de Busca

#### 1. CPF (EXATO)

```
✓ Busca por 11 dígitos
✓ Ignora formatação (. e -)
✓ Match exato

Exemplo:
  Entrada: "123.456.789-00"
  Busca: "12345678900"
  Resultado: Encontra registros com CPF = "12345678900"
```

#### 2. Email (VALIDAÇÃO + BUSCA)

```
✓ Validação: deve ter @ e .
✓ Busca: case-insensitive, com trim
✓ Match: coluna Email

Exemplo:
  Entrada: "JOAO@EMAIL.COM"
  Validação: ✓ Válido
  Busca: "joao@email.com"
  Resultado: Encontra registros com Email = "joao@email.com"
```

#### 3. Nome Completo (FUZZY)

```
✓ Validação: mínimo 2 partes, mínimo 3 caracteres
✓ Busca: case-insensitive, sem acentos
✓ Match: coluna NomeCompleto normalizado

Exemplo:
  Entrada: "João Silva"
  Validação: ✓ 2 partes, 11 caracteres
  Busca: "JOAO SILVA"
  Resultado: Encontra registros com NomeCompleto = "JOAO SILVA"
```

---

### 2.3 Regras de Validação de Entrada

#### 1. Lista de Itens (VALIDAÇÃO MÚLTIPLA)

```python
def validar_lista(itens: list) -> tuple:
    cpfs_validos = []
    emails_validos = []
    nomes_validos = []
    duplicatas = []
    
    for item in itens:
        # Validação de CPF
        if eh_cpf_valido(item):  # 11 dígitos
            if item in cpfs_validos:
                duplicatas.append(item)
            cpfs_validos.append(item)
        
        # Validação de Email
        elif eh_email_valido(item):  # @ + .
            emails_validos.append(item)
        
        # Validação de Nome Completo
        elif eh_nome_valido(item):  # 2+ partes, 3+ chars
            nomes_validos.append(item)
    
    return cpfs_validos, emails_validos, nomes_validos, duplicatas
```

#### 2. Detecção de Duplicatas

```
✓ CPFs duplicados na lista de entrada são identificados
⚠️ Aviso: "CPF duplicado: 123.456.789-00"
→ Apenas primeira ocorrência é processada
```

---

### 2.4 Fluxo de Inativação

#### Fase 1: BUSCAR (Preview)

```
[1] Upload de base + lista de itens

[2] Validação:
    - Validar extensão do arquivo
    - Validar formato da lista (CPF, Email ou Nome)
    - Detectar duplicatas

[3] Busca:
    - Por CPF: match exato
    - Por Email: match case-insensitive
    - Por Nome: match normalizado (sem acentos)

[4] Resultado:
    - Lista de usuários encontrados
    - Status: encontrado ou não encontrado
    - Dados: Id, Nome, Email, Status, Departamento, Empresa
    - Ordenação: encontrados primeiro, depois por nome
```

#### Fase 2: VALIDAR (Preview de Inativação)

```
[1] Usuário revisita encontrados

[2] Identificação de impacto:
    - Estruturas de aprovação afetadas
    - Estruturas que ficarão sem aprovador
    - Dados relacionados que serão removidos

[3] Relatório de impacto:
    - Quantidade de estruturas
    - Quantidade de aprovações
    - Status atual vs. pós-inativação

[4] Confirmação:
    - Usuário decide prosseguir ou cancelar
```

#### Fase 3: INATIVAR (Execução)

```
[1] Processar inativação:
    - Marcar usuário como inativo
    - Remover de estruturas de aprovação
    - Atualizar histórico

[2] Remover aprovações:
    - Se era único aprovador: marca estrutura como "sem aprovador"
    - Se havia segundo nível: promove segundo para primeiro
    - Se não há segundo nível: estrutura fica vazia

[3] Gerar saída:
    - Arquivo Excel com usuários inativados
    - Arquivo Excel com estruturas afetadas
    - Relatório de mudanças
```

---

### 2.5 Estrutura de Dados Retornada

#### Buscar Response

```python
{
    "total_buscados": 3,
    "total_encontrados": 2,
    "cpfs_duplicados": ["123.456.789-00"],
    "resultados": [
        {
            "id": "USR001",
            "nome": "João Silva",
            "cpf": "123.456.789-00",
            "email": "joao@empresa.com",
            "status": "Ativo",
            "departamento": "TI",
            "empresa": "Empresa A",
            "found": true,
            "tipo_busca": "cpf",
        },
        {
            "id": None,
            "nome": "Maria Santos",
            "cpf": "998.765.432-10",
            "email": "maria@empresa.com",
            "status": None,
            "departamento": None,
            "empresa": None,
            "found": false,
            "tipo_busca": "cpf",
        },
    ],
}
```

#### Estruturas Sem Aprovador

```python
"estruturas_sem_aprovador": [
    {
        "aprovacao_id": "APR001",
        "aprovacao_por": "VIAJANTE",
        "descricao": "Aprovação para reembolsos",
        "motivo": "Será removido único aprovador"
    }
]
```

---

### 2.6 Normalização de Colunas da Lista

O sistema tenta normalizar colunas de entrada:

```python
def _normalize_lista_columns(df_lista: pd.DataFrame):
    # Procura por CPF em qualquer coluna
    cpf_col = buscar_coluna_com("CPF")
    
    # Procura por Nome Completo em:
    # - "NOME COMPLETO"
    # - "NOME" + "SOBRENOME"
    nome_col = buscar_coluna_com("NOME COMPLETO")
    or
    nome_col = buscar_coluna_com("NOME") + buscar_coluna_com("SOBRENOME")
    
    # Procura por Email
    email_col = buscar_coluna_com("EMAIL")
    
    # Retorna DataFrame normalizado com:
    # - Coluna "CPF"
    # - Coluna "NomeCompleto"
    # - Coluna "Email"
```

---

### 2.7 Casos de Erro

| Situação | Erro | Resolução |
|----------|------|-----------|
| Arquivo vazio | "Envie a base (arquivo Excel)" | Fornecer arquivo |
| Extensão inválida | "Extensão não permitida. Aceitos: .xlsx, .xls, .xltx" | Usar formato correto |
| Lista vazia | "Nenhum item para buscar" | Fornecer lista de CPF/Email/Nome |
| CPF incompleto | CPF ignorado | Fornecer 11 dígitos |
| Email inválido | Email ignorado | Fornecer email válido |
| Nome muito curto | Nome ignorado | Fornecer 2+ partes, 3+ chars |
| Base sem CPF | "Base não contém coluna 'CPF'" | Adicionar coluna CPF |
| Nenhum resultado | Status "not found" | Verificar se dados existem |

---

## 3. MATRIZ COMPARATIVA

### Aprovação vs Inativação

| Aspecto | Aprovação | Inativação |
|---------|-----------|-----------|
| **Função** | Substituir/remover aprovadores | Buscar/inativar usuários |
| **Entrada** | Base com estruturas + CPF | Base com usuários + lista |
| **Critério** | AprovacaoId | CPF, Email, Nome |
| **Validação** | CPF na base de usuários | CPF/Email/Nome válido |
| **Saída** | Estruturas afetadas | Usuários inativados |
| **Impacto** | Modifica fluxos | Altera status de usuário |
| **Fallback** | Segundo nível (se houver) | N/A |

---

## 4. INTEGRAÇÃO COM CADASTRO

### Fluxo Completo de Usuário

```
[1] CADASTRO
    - Cria novo usuário
    - Valida 5 aspectos
    - Gera saída padrão
    ↓
[2] APROVAÇÃO
    - Estruturas de aprovação já existem previamente (cadastradas na Argo)
    - Substitui aprovadores quando alguém assume o lugar de outro
    ↓
[3] INATIVAÇÃO
    - Remove de estruturas
    - Marca como inativo
```

---

## 5. CASOS DE USO

### Caso 1: Substituir Aprovador

```
Entrada:
- CPF atual: 123.456.789-00
- CPF novo: 987.654.321-00
- Base: estruturas_viagem.xlsx
- Usuários: base_usuarios.xlsx

Processamento:
- Valida os 2 CPFs (11 dígitos, ATIVO na base)
- Localiza estruturas onde o CPF atual aparece
- Verifica se o CPF novo já aparece nelas (duplicidade)
- Substitui o CPF atual pelo novo, na mesma posição

Saída:
- 25 estruturas afetadas
- 42 posições substituídas
- 2 estruturas com duplicidade (novo já era aprovador)
```

### Caso 2: Remover Aprovador

```
Entrada:
- CPF: 123.456.789-00

Preview:
- 25 estruturas usam esse aprovador
- 5 ficarão sem aprovador
- 3 têm segundo nível (será promovido)

Execução:
- Remove de 25 estruturas
- Promove segundo nível em 3
- Marca 5 como "sem aprovador"

Saída:
- Arquivo com estruturas atualizadas
- Relatório de mudanças
```

### Caso 3: Inativar Usuário

```
Entrada:
- Lista: ["123.456.789-00", "maria@empresa.com", "João Silva"]

Busca:
- CPF 123.456.789-00: Encontrado ✓
- maria@empresa.com: Encontrado ✓
- João Silva: Não encontrado ✗

Inativação:
- Remove de 10 estruturas de aprovação
- Marca como inativo
- Gera relatório

Saída:
- 2 usuários inativados
- 10 estruturas afetadas
- 1 estrutura sem aprovador
```

---

## 6. Resumo Arquitetural

| Aspecto | Detalhe |
|---------|---------|
| **Modelos** | AprovacaoId, LoginAprovador_X, AprovacaoPor |
| **Validações** | CPF (11 dig), Base (colunas obrigatórias), Estruturas |
| **Fluxos** | Preview → Execução → Export |
| **Integração** | Com Cadastro e Inativação |
| **Saída** | Excel com dados modificados + relatórios |

