# ProcessData

Sistema web para **análise, cadastro e inativação de usuários a partir de planilhas Excel**, desenvolvido para automatizar rotinas operacionais e reduzir atividades manuais repetitivas.

O projeto foi criado a partir de uma **necessidade real do dia a dia profissional**, onde processos manuais em planilhas demandavam tempo, atenção constante e estavam sujeitos a erros.

---

## 🎯 Objetivo

- Automatizar processos manuais baseados em planilhas
- Padronizar validações e regras de negócio
- Reduzir retrabalho e falhas operacionais
- Centralizar fluxos de cadastro e inativação
- Servir como base organizada para evolução futura

---

## 🚀 Funcionalidades

- Upload e processamento de planilhas Excel
- Validação de colunas e campos obrigatórios
- Cadastro de usuários em massa
- Inativação de usuários a partir de base do cliente
- Geração de planilhas finais padronizadas, prontas para carga na plataforma Argo
- Interface web com abas de Análise, Cadastro, Inativação e Histórico
- Backend preparado para execução local e deploy

---

## 🛠️ Tecnologias Utilizadas

### Backend
- Python
- Flask
- Flask-CORS
- Pandas
- OpenPyXL
- xlrd

### Frontend
- HTML
- CSS
- JavaScript

---

## 📂 Estrutura do Projeto

```
ProcessData/
├─ backend/
│  ├─ api/            # blueprints Flask (cadastro, inativacao, aprovacao, analysis, history, frontend, health)
│  ├─ core/           # config e logging
│  ├─ domain/         # regras e contratos (MODEL_COLS, FICHA_MAP, REQUIRED_OUTPUT_COLS)
│  ├─ services/       # casos de uso (processing, validation, inactivation, export, audit, report)
│  ├─ shared/         # utilitarios (texto, cpf, arquivos, validacao de upload)
│  ├─ infra/          # persistencia (trilha de auditoria JSONL)
│  ├─ processor.py    # motor de inativacao (processar_inativacao_from_paths)
│  ├─ app.py          # factory Flask
│  └─ tests/          # suite pytest
│
├─ frontend/
│  ├─ index.html
│  └─ static/{css,js}/   # js: app.v2.js + modulos por aba (analise, inativacao, history)
│
├─ tests/e2e/         # Playwright (fluxos criticos)
├─ pyproject.toml
├─ requirements.txt
├─ .gitignore
└─ README.md
```

---

## ⚡ Quick Start

### Pré-requisitos
- Python 3.10+
- Git
- (opcional) Node 18+ para os testes end-to-end Playwright

### Executar localmente

```bash
git clone https://github.com/alweedev/ProcessData.git
cd ProcessData
python -m venv .venv
# Windows
.\.venv\Scripts\Activate.ps1
# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt   # requirements.txt fica na raiz
python -m backend.app             # ou: python backend/app.py
```

### Variáveis de ambiente (opcionais)

| Var | Padrão | Uso |
|-----|--------|-----|
| `UPLOAD_FOLDER` | `<tmp>/processdata_uploads` | pasta de arquivos temporários (fora do repo) |
| `HISTORY_LOG_FILE` | `<UPLOAD_FOLDER>/history.log.jsonl` | trilha de auditoria (JSONL, rotacionada) |
| `HISTORY_MAX_BYTES` | `5242880` | tamanho para rotacionar a auditoria |
| `CORS_ORIGINS` | *(vazio = mesma origem)* | lista separada por vírgula de origens permitidas em `/api/*` |
| `HISTORY_ADMIN_TOKEN` | *(vazio)* | token para `DELETE /api/history` fora de localhost |
| `HOST` / `PORT` / `DEBUG` | `0.0.0.0` / `5000` / `false` | servidor Flask |

## Acesse no Navegador

```
http://127.0.0.1:5000
```

## 🔁 Fluxos Principais

### Cadastro em massa

1. Acesse a aba **Cadastro**.
2. Faça upload da ficha de cadastro no formato esperado.
3. O sistema:
   - Normaliza colunas conforme o modelo definido
   - Separa **Nome** e **Sobrenome** a partir de **Nome Completo**
   - Valida CPF, e-mail e campos obrigatórios
   - Gera a planilha final pronta para uso

---

### Inativação de usuários

1. Acesse a aba **Inativação**.
2. Faça upload da base de usuários do cliente.
3. Informe a lista de desligados (CPF, nome completo ou e-mail).
4. O sistema:
   - Realiza match por CPF, nome e e-mail
   - Preserva **Nome** e **Sobrenome** da base original
   - Gera a planilha final de inativação

## 🧪 Testes

Com o ambiente virtual ativo, a partir da raiz do repositório:

```bash
python -m pytest -q                    # suíte completa (backend/tests/)
python -m pytest backend/tests/test_cadastro_api.py -q
python -m pytest backend/tests/test_inativacao_api.py backend/tests/test_inativacao_generation_api.py -q
python -m pytest backend/tests/test_aprovacao.py -q
```

Fluxos críticos ponta-a-ponta (Playwright, requer Node):

```bash
npm ci
npx playwright install
npx playwright test
```

## 🔒 Segurança

- **Excel**: células que começam com `= + - @` / TAB / CR são gravadas como
  texto (anti *formula injection*).
- **Preview**: valores da planilha são escapados antes de ir para o DOM.
- **Upload**: além da extensão, o conteúdo `.xlsx/.xltx` é checado (assinatura
  ZIP + estrutura OOXML).
- **Estático**: o servidor só entrega arquivos dentro de `frontend/` (bloqueia `../`).
- **CORS**: mesma origem por padrão; libere origens com `CORS_ORIGINS`.
- **Histórico**: `DELETE /api/history` só de localhost ou com `X-Admin-Token`.
- **Auditoria**: `history.log.jsonl` rotaciona ao passar `HISTORY_MAX_BYTES`.
- **Erros**: respostas 5xx são genéricas; o detalhe fica só no log do servidor.
- Não versione planilhas reais, `.env`, `.venv`, logs sensíveis ou temporários
  (`UPLOAD_FOLDER` fica fora do repositório por padrão).

## 📌 Observações

- O projeto segue em evolução contínua com foco em organização, clareza e boas práticas.
- Ferramentas de IA foram utilizadas como suporte ao desenvolvimento, principalmente para revisão de código, identificação de melhorias e aceleração do aprendizado, com todas as decisões técnicas sendo analisadas e implementadas conscientemente.

- ## 👨‍💻 Autor

Desenvolvido por **Alejandro Gabriel**

LinkedIn: https://www.linkedin.com/in/alejandro-gabriel/

