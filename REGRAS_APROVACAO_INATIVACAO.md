# Regras de Negócio — Aprovação e Inativação

Cobre as duas abas que operam sobre a base de usuários/estruturas já
cadastrada na Argo (diferente do cadastro, que **cria** registros — ver
`REGRAS_NEGOCIO_PROCESSAMENTO.md`).

| Aba | Código | Endpoints |
|---|---|---|
| Estruturas de Aprovação | `backend/services/approval_service.py` (`ApprovalService`) + `backend/api/aprovacao.py` | `/api/aprovacao/remover/{preview,export}`, `/api/aprovacao/substituir/{preview,export}` |
| Inativação | `backend/services/inactivation_service.py` (`InactivationService`) + `backend/processor.py` (`processar_inativacao_from_paths`) + `backend/api/inativacao.py` | `/api/inativacao/buscar`, `/api/preview_inativacao`, `/api/process_inativacao` |

> **As duas abas são independentes.** Inativar um usuário **não** remove os
> aprovadores dele das estruturas de aprovação automaticamente — isso é uma
> ação manual separada, feita depois na aba Estruturas (Remover Aprovador),
> se necessário. Não existe hoje nenhuma chamada entre os dois módulos.

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
> módulo 11); cadastro e inativação validam só o comprimento (11 dígitos).
> A base de usuários precisa ter coluna `CPF` e uma de nome (`NomeCompleto`
> ou `Nome` [+ `SobreNome`]) — senão: `"Base de usuários não contém coluna
> 'CPF'."` / `"...coluna de nome ('NomeCompleto' ou 'Nome')."`.

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

## 2. Inativação

### 2.1 Objetivo

Buscar usuários na base do cliente (por CPF, e-mail ou nome completo) e
gerar a ficha de inativação (`Operacao=DELETE`) pronta para carga na Argo.
**Não** altera a base de origem nem mexe em estruturas de aprovação — só
produz o arquivo de saída.

### 2.2 Critérios de busca (`InactivationService.search_matches`)

| Critério | Regra |
|---|---|
| CPF | 11 dígitos após limpar pontuação; match exato |
| Email | precisa casar `^[^@\s]+@[^@\s]+\.[^@\s]+$`; match case-insensitive |
| Nome completo | mínimo 2 partes e 3 caracteres; match normalizado (maiúsculas, sem acento) |

A busca roda nas três frentes ao mesmo tempo sobre a lista de itens
informada (colados como texto, um por linha, ou upload de planilha). CPFs
repetidos na lista de entrada são reportados em `duplicates`.

Resposta de `/api/inativacao/buscar`:

```jsonc
{
  "items": [
    { "id": null, "nome": "João Silva", "cpf": "12345678900",
      "email": "joao@empresa.com", "status_atual": "ATIVO", "found": true },
    { "id": null, "nome": "", "cpf": "99876543210", "email": "",
      "status_atual": "Não localizado", "found": false }
  ],
  "total": 2,
  "duplicates": [],
  "not_found": ["99876543210"]
}
```

> `id` só vem preenchido se a base tiver uma coluna reconhecível como
> `UserId` (ver `_detect_base_cols`) — na maioria das bases fica `null`.

Resultados vêm ordenados: encontrados primeiro, depois por nome/CPF.

### 2.3 Geração da ficha (`processar_inativacao_from_paths`)

```
[1] Filtra a base para Status = ATIVO (quando a coluna existe)
[2] Casa a lista contra a base, nessa ordem de prioridade:
    CPF (exato) → NomeCompleto (exato, normalizado) → Email (exato,
    case-insensitive) — cada usuário conta uma vez só (a 1ª forma de match
    que encontrar ele "ganha")
[3] Monta a saída: Operacao="DELETE", campos copiados da base
    (UserId, Login, Nome, SobreNome, Email, Cargo, Departamento, Nivel,
    NomeEmpresa, CC, flags booleanas), EmpresaCCustoParaUsuario="S",
    CodigoIntegracao="AUT"
```

`/api/preview_inativacao` roda o mesmo processamento e devolve uma amostra
(10 linhas) e os registros completos (até 500) para conferência antes do
download; `/api/process_inativacao` gera o `.xlsx` final para download.

### 2.4 Casos de erro

| Situação | Erro |
|---|---|
| Sem base | `Envie a base (arquivo Excel)` |
| Extensão inválida | `Extensão não permitida. Aceitos: .xlsx, .xls, .xltx` |
| Sem lista (nem arquivo, nem texto colado) | `Envie a lista como arquivo ou cole nomes/CPFs no campo de texto` |
| Texto colado sem CPF/nome/e-mail válido | `Texto de lista vazio ou sem CPF/Nome/E-mail válidos` |
| Base sem coluna CPF/Nome/Email reconhecível | busca correspondente fica vazia para esse critério (sem erro explícito) |

---

## 3. Comparativo

| Aspecto | Aprovação | Inativação |
|---|---|---|
| Ação | Substitui/remove aprovadores em estruturas existentes | Gera ficha de saída (DELETE) para usuários encontrados |
| Entrada | Base de estruturas + base de usuários + CPF(s) | Base de usuários + lista (CPF/e-mail/nome) |
| Chave de match | `AprovacaoId` / CPF do aprovador | CPF → Nome → Email (nessa prioridade) |
| Valida dígito verificador de CPF | Sim | Não (só comprimento) |
| Saída | Base de estruturas atualizada (`Operacao=UPDATE` nas linhas alteradas) | Ficha de inativação (`Operacao=DELETE`) |
| Efeito colateral automático no outro módulo | Nenhum | Nenhum |
