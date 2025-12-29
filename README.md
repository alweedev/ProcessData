# Robo de Cadastro / Inativação

Sistema web para análise, cadastro e inativação de usuários a partir de planilhas Excel, com backend em Python/Flask e frontend em HTML/CSS/JS.

---

## Estrutura do projeto

Raiz do repositório:

- `backend/` – API Flask, regras de negócio e validações.
- `frontend/` – interface web (tela com abas: Análise, Cadastro, Inativação, Histórico).
- `data/` – espaço sugerido para planilhas de exemplo (entrada/saída) que você queira versionar.
- `tmp_uploads/` – pasta de trabalho para uploads temporários (não deve ser versionada em produção).
- `README.md` – esta documentação.

### Backend (`backend/`)

Arquivos principais:

- `app.py` – ponto de entrada Flask. Cria a aplicação, registra os blueprints de API e frontend e sobe o servidor.
- `api/cadastro.py` – rotas de cadastro em lote (upload de ficha, processamento e geração de planilha de saída).
- `api/inativacao.py` – rotas de inativação (upload de base do cliente, lista de desligados, preview e planilha final).
- `api/frontend.py` – rota que serve o frontend (HTML inicial).
- `core/config.py` – configurações gerais (paths, flags de debug, etc.).
- `core/logging.py` – configuração de logging (formato e destino dos logs).
- `processor.py` – regras de negócio de cadastro e inativação (leitura de arquivos, normalização de colunas, montagem da ficha modelo, etc.).
- `utils.py` – funções utilitárias compartilhadas (normalização de texto, CPF, separação de nome/sobrenome, etc.).
- `validators.py` – validações de linha e do DataFrame final (colunas obrigatórias, CPF, e-mail, nível, etc.).
- `requirements.txt` – dependências Python do backend.

Scripts de teste auxiliar:

- `test_cadastro_run.py` – fluxo de teste rápido do cadastro.
- `test_inativacao_run.py` – fluxo de teste rápido da inativação.
- `test_integration_api.py` – teste simples de integração de algumas rotas da API.

> Observação: arquivos temporários de teste (por exemplo, planilhas geradas manualmente em `backend/`) não são necessários para entrega em produção. Use a pasta `data/` ou remova-os antes de entregar o projeto.

### Frontend (`frontend/`)

Arquivos principais:

- `index.html` – página principal com as abas de Análise, Cadastro, Inativação e Histórico.
- `static/css/` – estilos do projeto (`custom.css`, animações, etc.), incluindo suporte a tema claro/escuro.
- `static/js/app.v2.js` – script principal da aplicação (controle de abas, tema, histórico, ajuda, etc.).
- `static/js/inativacao/` – módulo dedicado ao fluxo de inativação:
  - `state.js` – estado da tela (lista de itens, base carregada, resultados).
  - `api.js` – chamadas às rotas de inativação no backend.
  - `validation.js` – validação de lista (CPF, nome, e-mail).
  - `ui.js` – renderização de tabela, contadores e indicadores visuais.
  - `controller.js` – orquestração dos eventos da tela.

> Importante: o frontend é servido pelo próprio backend Flask (via blueprint de frontend). Na execução local padrão, acesse `http://127.0.0.1:5000` no navegador.

---

## Requisitos

- Python 3.10+ (recomendado; versões a partir de 3.8 tendem a funcionar).
- Ambiente virtual (`venv`) para isolar dependências.

---

## Como rodar localmente

1. **Criar e ativar o ambiente virtual** (Windows PowerShell, na raiz do projeto):

   ```powershell
   cd c:\novo_script2
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. **Instalar as dependências do backend**:

   ```powershell
   cd backend
   pip install -r requirements.txt
   ```

3. **Subir o servidor Flask**:

   ```powershell
   cd backend
   python app.py
   ```

4. **Acessar o sistema no navegador**:

   - URL padrão (ambiente local): `http://127.0.0.1:5000`

---

## Fluxos principais

### Cadastro em lote

1. Acesse a aba **Cadastro** na interface.
2. Faça o upload da ficha de cadastro no formato esperado (colunas mapeadas conforme `FICHA_MAP` em `backend/processor.py`).
3. O backend:
   - Normaliza colunas conforme o modelo (`MODEL_COLS`).
   - Separa `Nome` e `SobreNome` a partir de `NomeCompleto` usando regras de nomes compostos / sobrenomes compostos.
   - Valida CPF, e-mail, campos obrigatórios e níveis.
   - Gera uma planilha de saída no layout padrão para importação no sistema destino.

### Inativação

1. Acesse a aba **Inativação**.
2. Faça o upload da **base de usuários do cliente**.
3. Informe a lista de desligados (CPF, Nome Completo ou e-mail), via textarea ou arquivo conforme configurado.
4. O backend:
   - Faz match por CPF, NomeCompleto e e-mail.
   - Respeita o `Nome`/`SobreNome` da base do cliente (não recalcula esses campos na saída de inativação).
   - Gera uma planilha de inativação com os usuários encontrados.

---

## Executando testes rápidos

Com o ambiente virtual ativo e as dependências instaladas, você pode rodar os scripts de teste auxiliares no backend:

```powershell
cd backend
python test_cadastro_run.py
python test_inativacao_run.py
python test_integration_api.py
```

Esses testes fazem verificações rápidas dos fluxos de cadastro, inativação e algumas rotas da API.

---

## Deploy na Railway

Este projeto está preparado para deploy na Railway usando Nixpacks e Gunicorn.

### O que já está configurado

- Procfile na raiz com comando web: `gunicorn backend.app:app --workers 2 --threads 8 --bind 0.0.0.0:$PORT`.
- railway.toml com `start` equivalente (opcional; a Railway pode usar o Procfile).
- requirements.txt inclui `gunicorn` e dependências do Flask.
- O backend lê `PORT`, `HOST` e `DEBUG` de variáveis de ambiente (veja `backend/core/config.py`).

### Passo a passo

1. Crie um projeto na Railway e conecte este repositório (GitHub).
2. A Railway detectará Python e instalará as dependências via `requirements.txt`.
3. O serviço web iniciará com Gunicorn usando o Procfile.
4. A URL pública servirá tanto a API quanto o frontend (o blueprint `frontend` entrega `frontend/index.html`).

### Variáveis de ambiente

- PORT: definida pela Railway automaticamente.
- HOST: opcional (padrão `0.0.0.0`).
- DEBUG: opcional (`true`/`false`). Em produção, deixe `false`.

### Health check

- Endpoint: `/health` (retorna 200 com `{status: "OK"}` ou 204 em HEAD).

### Teste local com Gunicorn (opcional)

Com o `venv` ativo na raiz:

```powershell
pip install -r requirements.txt
gunicorn backend.app:app --workers 2 --threads 8 --bind 0.0.0.0:5000
```

Abra no navegador: `http://127.0.0.1:5000`.

---

## Problemas comuns

- **405 Method Not Allowed**
  - Verifique se o backend está rodando na porta `5000`.
  - Confirme se as rotas `/api/process_cadastro` e `/api/process_inativacao` estão sendo chamadas com método **POST**.

- **Erros de dependência (ImportError / ModuleNotFoundError)**
  - Certifique-se de que os pacotes do `requirements.txt` foram instalados no mesmo `venv` que o VS Code / terminal está usando.

- **Frontend não atualiza ou carrega scripts antigos**
  - O frontend utiliza `frontend/static/js/app.v2.js` e o módulo `frontend/static/js/inativacao/`.
  - O arquivo antigo `app.js` não é mais utilizado.
  - Se mudanças não aparecerem, tente limpar o cache do navegador (Ctrl+F5).

---

## Observações finais

- A pasta `tmp_uploads/` é usada apenas para arquivos temporários enviados pelo usuário em tempo de execução e não deve conter dados sensíveis versionados.
- Antes de entregar o projeto à empresa, recomenda-se remover arquivos de teste manuais (como planilhas geradas para depuração) ou movê-los para `data/` apenas como exemplos documentados.
