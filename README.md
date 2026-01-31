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
- python-docx

### Frontend
- HTML
- CSS
- JavaScript

---

## 📂 Estrutura do Projeto

```
ProcessData/
├─ backend/
│  ├─ api/
│  │  ├─ cadastro.py
│  │  ├─ inativacao.py
│  │  └─ frontend.py
│  ├─ core/
│  │  ├─ config.py
│  │  └─ logging.py
│  ├─ processor.py
│  ├─ utils.py
│  ├─ validators.py
│  ├─ test_cadastro_run.py
│  ├─ test_inativacao_run.py
│  └─ test_integration_api.py
│
├─ frontend/
│  ├─ index.html
│  └─ static/
│     ├─ css/
│     └─ js/
│        └─ inativacao/
│
├─ data/          # (opcional) exemplos de planilhas fictícias
├─ tmp_uploads/   # pasta temporária (não versionada)
├─ .gitignore
├─ Procfile
├─ railway.toml
└─ README.md
```

---

## ⚡ Quick Start

### Pré-requisitos
- Python 3.8+
- Git

### Executar localmente

```
bash
git clone https://github.com/alweedev/ProcessData.git
cd ProcessData
python -m venv .venv
# Windows
.\.venv\Scripts\Activate.ps1
# Linux / macOS
source .venv/bin/activate

cd backend
pip install -r requirements.txt
python app.py

```

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

## 🧪 Testes Rápidos

Com o ambiente virtual ativo:

```
bash
cd backend
python test_cadastro_run.py
python test_inativacao_run.py
python test_integration_api.py
```

## 📌 Observações

- A pasta `tmp_uploads/` é utilizada apenas em tempo de execução e não deve conter dados sensíveis.
- Planilhas reais não devem ser versionadas. Utilize apenas exemplos fictícios em `data/`.
- O projeto segue em evolução contínua com foco em organização, clareza e boas práticas.
- Ferramentas de IA foram utilizadas como suporte ao desenvolvimento, principalmente para revisão de código, identificação de melhorias e aceleração do aprendizado, com todas as decisões técnicas sendo analisadas e implementadas conscientemente.

- ## 👨‍💻 Autor

Desenvolvido por **Alejandro Gabriel**

LinkedIn: https://www.linkedin.com/in/alejandro-gabriel/

