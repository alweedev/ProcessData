# Arquitetura Modernizada - ProcessData

## Objetivo

Arquitetura modular, legível e testável, preservando as regras de negócio de
cadastro, inativação e aprovação. Para o detalhe de cada regra, ver
`REGRAS_NEGOCIO_PROCESSAMENTO.md` (cadastro) e
`REGRAS_APROVACAO_INATIVACAO.md` (aprovação/inativação); o `README.md` traz
quick start e visão geral.

## Camadas

- `backend/domain`: regras e contratos (`MODEL_COLS`, `FICHA_MAP`,
  `REQUIRED_OUTPUT_COLS`) — **fonte única** desses dados.
- `backend/services`: orquestração de casos de uso (`ProcessingService`,
  `ValidationService`, `InactivationService`, `ApprovalService`,
  `ExportService`, `AuditService`, `ReportService`).
- `backend/shared`: utilitários compartilhados — `text_utils` (normalização),
  `cpf_utils` (limpeza, formatação, dígito verificador), `file_utils` (extensão,
  nome temporário), `upload_validation` (sniff de conteúdo).
- `backend/infra`: integrações técnicas (trilha de auditoria JSONL com rotação).
- `backend/api`: rotas HTTP e serialização.
- `backend/processor.py`: motor de **inativação** (`processar_inativacao_from_paths`).

## Estado da consolidação (set/2026)

- O pipeline de **cadastro** roda exclusivamente por
  `ProcessingService.process_records_from_files`. A versão legada
  (`processor.processar_registros_from_files`) e `backend/validators.py` foram
  removidas; os comportamentos válidos (validação `__geral__`, `Terceiro` com
  dígitos, `DescricaoCCustoEmpresa` só sem acento, guarda de NaN no prefixo
  FRONT, fold de acento no S/N) foram portados para os serviços com teste.
- `backend/utils.py` foi consolidado em `backend/shared/*`.
- `MODEL_COLS`/`FICHA_MAP` deixam de ter cópias divergentes: `processor.py`
  importa de `backend.domain.rules`.
- Suporte a `.docx` foi removido (função `extrair_docx` nunca existiu). Entrada:
  `.xlsx`, `.xls`, `.xltx`.
- A lógica de **aprovação** (normalização de CPF, detecção de colunas, cálculo
  de impacto, remoção/compactação) saiu de `api/aprovacao.py` pra
  `ApprovalService`, no mesmo padrão de `ExportService`/`InactivationService`.
- Fase 8: fluxo de **substituição de aprovador** (`ApprovalService.replace_cpf`
  / `check_new_approver_duplicates`), com gate de duplicidade quando o novo
  CPF já é aprovador na estrutura. Ver `REGRAS_APROVACAO_INATIVACAO.md` §1.4.

## Endpoints

| Método | Rota | Uso |
|---|---|---|
| POST | `/api/process_cadastro` | gera a ficha de cadastro (xlsx) |
| POST | `/api/analysis/summary` | relatório de qualidade + preview (JSON) |
| POST | `/api/inativacao/buscar` | busca (encontrado/não encontrado) por CPF/nome/e-mail |
| POST | `/api/preview_inativacao` | preview da ficha de inativação antes da geração |
| POST | `/api/process_inativacao` | gera a ficha de inativação (xlsx) |
| POST | `/api/aprovacao/remover/preview` | impacto da remoção de um aprovador |
| POST | `/api/aprovacao/remover/export` | base de aprovação atualizada (xlsx) |
| POST | `/api/aprovacao/substituir/preview` | impacto da substituição de um aprovador por outro |
| POST | `/api/aprovacao/substituir/export` | base de aprovação atualizada após substituição (xlsx) |
| GET/DELETE | `/api/history` | trilha de auditoria (DELETE só localhost/token) |
| GET | `/api/health` | health check |

Removidos: `POST /api/inativacao/executar` (sucesso falso, sem consumidor) e
`GET /health` (duplicava `/api/health`).

## Segurança

- **Formula injection**: `ExportService` grava como texto explícito qualquer
  célula que comece com `= + - @` / TAB / CR.
- **XSS**: o preview de inativação escapa as células vindas da planilha.
- **Uploads**: além da extensão, `validar_conteudo_xlsx` verifica assinatura
  ZIP + estrutura OOXML (`.xls` fica a cargo do pandas/xlrd).
- **Path traversal**: o servidor estático só entrega arquivos cujo caminho real
  resolve para dentro de `frontend/`.
- **CORS**: restrito à mesma origem por padrão; `CORS_ORIGINS` (env) libera
  origens específicas.
- **Histórico**: `DELETE /api/history` exige localhost ou `X-Admin-Token`.
- **Auditoria**: JSONL rotacionado por tamanho (`HISTORY_MAX_BYTES`); leitura
  limitada à cauda.
- **Logs/erros**: respostas 5xx genéricas ao cliente (detalhe só no log do
  servidor); logs não incluem nome/colunas.

## Testes

`python -m pytest -q --cov` (suíte em `backend/tests/`, 123/123, cobertura
medida mas sem gate de threshold ainda). `npm --prefix frontend-react run
test` (Vitest, funções puras dos hooks). Fluxos críticos ponta-a-ponta em
`tests/e2e/` (Playwright, 19 specs).

## Próxima fase

- Sem migração de framework planejada no curto prazo: a stack atual (Flask +
  services) está estável, testada (123/123 pytest, Vitest e Playwright
  verdes) e sem sinal de dor de crescimento no código. Avaliar FastAPI +
  Pydantic v2 + SQLAlchemy 2 + Alembic + PostgreSQL fica registrado, mas
  parado até existir um driver concreto (ex: multi-tenant, autenticação,
  consulta SQL no histórico) — não faz sentido pagar esse custo pra um único
  cliente interno.
- Foco atual é manter a ferramenta sólida pro uso interno. A aba **Análise**
  foi removida na Fase 5 da migração do frontend, mas `/api/analysis/summary`
  **continua em uso** — virou a validação prévia (não-bloqueante) do
  Cadastro (`useCadastro.ts` chama `postAnalysisSummary` antes de gerar a
  ficha). Não há endpoint órfão aqui.
