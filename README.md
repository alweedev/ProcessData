# ProcessData

Sistema web para **cadastro em massa, inativação e gestão de estruturas de aprovação** a partir de planilhas Excel — automatiza rotinas operacionais que antes eram feitas manualmente em planilha, reduzindo retrabalho e erro humano.

Nasceu de uma necessidade real do dia a dia profissional: preparar fichas para carga na plataforma **Argo** (cadastro, inativação, aprovação) exigia validar e formatar planilhas manualmente, um processo repetitivo e sujeito a erro.

---

## 🚀 Funcionalidades

- **Cadastro em massa**: normaliza uma planilha de entrada (com nomes de coluna variados) para a ficha padrão de carga, com validação linha a linha e geral.
- **Inativação**: busca usuários numa base do cliente por CPF, e-mail ou nome e gera a ficha de desligamento (`Operacao=DELETE`).
- **Estruturas de aprovação**: substitui ou remove aprovadores em cadeias de aprovação existentes, com preview de impacto (estruturas afetadas, duplicidade, estrutura que ficaria sem aprovador) antes de exportar.
- Interface web em React, com abas de Início, Cadastro, Inativação, Estruturas e Histórico.
- Trilha de auditoria (JSONL, rotacionada) de toda exportação/ação sensível.

## 🛠️ Stack

| Camada | Tecnologias |
|---|---|
| Backend | Python, Flask, Flask-CORS, Pandas, OpenPyXL, xlrd |
| Frontend | React 19 + TypeScript, Vite, Tailwind CSS 4 |
| Testes | pytest (backend), Playwright (e2e), ruff + mypy (lint/tipos) |

## 📂 Estrutura do projeto

```
ProcessData/
├─ backend/
│  ├─ api/            # blueprints Flask (cadastro, inativacao, aprovacao, analysis, history, frontend, health)
│  ├─ core/           # config e logging
│  ├─ domain/         # regras e contratos (MODEL_COLS, FICHA_MAP, REQUIRED_OUTPUT_COLS)
│  ├─ services/       # casos de uso (processing, validation, inactivation, approval, export, audit, report)
│  ├─ shared/         # utilitários (texto, cpf, arquivos, validação de upload)
│  ├─ infra/          # persistência (trilha de auditoria JSONL)
│  ├─ processor.py    # motor de inativação (processar_inativacao_from_paths)
│  ├─ app.py          # factory Flask
│  └─ tests/          # suíte pytest
│
├─ frontend/
│  ├─ index.html         # shell servido pelo Flask; monta o bundle React
│  └─ static/react/      # build do frontend React (Vite) — gerado por `npm run build:react`
│
├─ frontend-react/    # projeto Vite + React + Tailwind (fonte da UI)
├─ tests/e2e/         # Playwright (fluxos críticos)
├─ data/              # fixtures/planilhas de teste manual (fora da suíte automatizada)
├─ pyproject.toml     # config de ruff/mypy/pytest
├─ requirements.txt
└─ package.json       # scripts de build/dev do frontend + testes e2e
```

---

## ⚡ Quick start

### Pré-requisitos

- Python 3.10+
- Node 18+ (build do frontend e testes e2e)
- Git

### Rodar o backend (serve o frontend já buildado)

```bash
git clone https://github.com/alweedev/ProcessData.git
cd ProcessData
python -m venv .venv
.\.venv\Scripts\Activate.ps1     # Windows
# source .venv/bin/activate      # Linux / macOS

pip install -r requirements.txt
python -m backend.app            # http://127.0.0.1:5000
```

### Desenvolver o frontend (com hot reload)

```bash
npm install
npm run dev:react     # Vite em http://127.0.0.1:5173, com HMR
```

Para gerar o build consumido pelo Flask (`frontend/static/react/`):

```bash
npm run build:react          # build único
npm run build:react:watch    # rebuild automático a cada mudança
```

### Variáveis de ambiente (opcionais)

| Var | Padrão | Uso |
|---|---|---|
| `UPLOAD_FOLDER` | `<tmp>/processdata_uploads` | pasta de arquivos temporários (fora do repo) |
| `HISTORY_LOG_FILE` | `<UPLOAD_FOLDER>/history.log.jsonl` | trilha de auditoria (JSONL, rotacionada) |
| `HISTORY_MAX_BYTES` | `5242880` (5 MB) | tamanho do arquivo de auditoria para rotacionar |
| `HISTORY_MAX_ROWS` | `5000` | linhas mantidas em memória ao ler o histórico |
| `HISTORY_ADMIN_TOKEN` | *(vazio)* | token para `DELETE /api/history` fora de localhost |
| `CORS_ORIGINS` | *(vazio = mesma origem)* | lista separada por vírgula de origens permitidas em `/api/*` |
| `HOST` / `PORT` / `DEBUG` | `0.0.0.0` / `5000` / `false` | servidor Flask |

Upload de arquivo é limitado a 16 MB (`MAX_CONTENT_LENGTH`, fixo).

---

## 🔁 Fluxos principais

### Cadastro em massa

1. Aba **Cadastro** → upload da planilha de entrada.
2. O sistema mapeia colunas conhecidas (nomes variados são aceitos, ver
   `FICHA_MAP`), separa **Nome**/**Sobrenome** a partir de **Nome Completo**,
   valida CPF/e-mail/campos obrigatórios e gera a ficha final.

### Inativação de usuários

1. Aba **Inativação** → upload da base de usuários do cliente + lista de
   desligados (CPF, nome completo ou e-mail, colados ou em planilha).
2. O sistema casa por CPF → nome → e-mail (nessa prioridade) e gera a ficha
   de inativação (`Operacao=DELETE`).

### Estruturas de aprovação

1. Aba **Estruturas** → upload da base de estruturas + base de usuários.
2. Informe o CPF do aprovador (e, no modo substituição, o CPF do novo
   aprovador).
3. Revise o preview de impacto (estruturas afetadas, avisos de duplicidade
   ou de estrutura que ficaria sem aprovador) e confirme a exportação — em
   todas as estruturas encontradas ou só nas selecionadas.

> Regras de negócio completas de cada fluxo:
> [`REGRAS_NEGOCIO_PROCESSAMENTO.md`](REGRAS_NEGOCIO_PROCESSAMENTO.md)
> (cadastro) e
> [`REGRAS_APROVACAO_INATIVACAO.md`](REGRAS_APROVACAO_INATIVACAO.md)
> (aprovação/inativação). Arquitetura, camadas e endpoints:
> [`ARQUITETURA_MODERNIZADA.md`](ARQUITETURA_MODERNIZADA.md).

---

## 🧪 Testes

Com o ambiente virtual ativo, a partir da raiz do repositório:

```bash
python -m pytest -q                              # suíte completa (backend/tests/), 116/116
python -m pytest backend/tests/test_cadastro_api.py -q
python -m pytest backend/tests/test_inativacao_api.py backend/tests/test_inativacao_buscar_api.py -q
python -m pytest backend/tests/test_aprovacao.py backend/tests/test_aprovacao_substituir.py -q
python -m pytest --cov --cov-report=term-missing  # cobertura (também roda no CI)

ruff check backend/                              # lint
mypy backend/                                    # tipos
```

Fluxos críticos ponta-a-ponta (Playwright, requer Node):

```bash
npm ci
npx playwright install
npm run test:e2e             # builda o frontend (npm run build:react) e roda as 17 specs
```

`npx playwright test` direto **não** rebuilda o frontend antes — use depois
de mudanças em `frontend-react/` ou prefira `npm run test:e2e`. O servidor
Flask do teste sobe via `python`; se não estiver no `PATH`, aponte pro
interpretador do venv com `PYTHON=.venv/Scripts/python.exe` (Windows) antes
do comando.

---

## 🔒 Segurança

- **Formula injection**: células que começam com `= + - @` / TAB / CR são
  gravadas como texto explícito na exportação.
- **XSS**: valores vindos da planilha são escapados no preview antes de ir
  para o DOM.
- **Upload**: além da extensão (`.xlsx`, `.xls`, `.xltx`), o conteúdo
  `.xlsx`/`.xltx` é checado (assinatura ZIP + estrutura OOXML); limite de
  16 MB por arquivo.
- **Path traversal**: o servidor estático só entrega arquivos cujo caminho
  real resolve para dentro de `frontend/`.
- **CORS**: restrito à mesma origem por padrão; `CORS_ORIGINS` libera
  origens específicas.
- **Histórico**: `DELETE /api/history` exige localhost ou `X-Admin-Token`;
  arquivo rotaciona por tamanho (`HISTORY_MAX_BYTES`).
- **Erros**: respostas 5xx são genéricas ao cliente; o detalhe fica só no
  log do servidor (sem vazar nome/colunas da planilha).
- Não versione planilhas reais, `.env`, `.venv` ou logs/temporários —
  `UPLOAD_FOLDER` já fica fora do repositório por padrão.

## 📌 Observações

- Projeto em evolução contínua, com foco em organização, clareza e boas
  práticas — sem migração de framework planejada no curto prazo (ver
  `ARQUITETURA_MODERNIZADA.md` § Próxima fase).
- Ferramentas de IA foram usadas como apoio ao desenvolvimento (revisão de
  código, identificação de melhorias, aceleração do aprendizado), com todas
  as decisões técnicas analisadas e implementadas conscientemente.

## 👨‍💻 Autor

Desenvolvido por **Alejandro Gabriel**

LinkedIn: https://www.linkedin.com/in/alejandro-gabriel/
