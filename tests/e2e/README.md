# Testes end-to-end (Playwright)

Cobrem os fluxos críticos na UI real (Chromium), contra o backend Flask.

## Pré-requisitos

- Python 3.10+ com as dependências instaladas (`pip install -r requirements.txt`)
- Node 18+

## Rodar

```bash
npm ci
npx playwright install --with-deps chromium
npx playwright test
```

O `webServer` do `playwright.config.ts` sobe `python -m backend.app` na porta
`E2E_PORT` (padrão 5001), com `UPLOAD_FOLDER`/`HISTORY_LOG_FILE` apontando para o
diretório temporário do SO. Para rodar contra um servidor já no ar, exporte
`E2E_PORT` correspondente — `reuseExistingServer` está ligado fora de CI.

## Specs

| Arquivo | Verifica |
|---|---|
| `cadastro-flow.spec.js` | upload + download `saida_cadastro.xlsx`, limpeza da seleção, guarda de 5 arquivos |
| `inativacao-flow.spec.js` | fluxo em 3 etapas: analisar → conferir impacto → confirmar/executar → baixa `inativacao.zip` (com segunda confirmação para estrutura órfã) |
| `preview-xss.spec.js` | célula maliciosa da planilha renderiza como texto (sem `alert`/execução) |

Fixtures são geradas em memória por `fixtures.mjs` (SheetJS) — não há `.xlsx`
versionado. Os seletores seguem os IDs de `frontend/index.html`; ajuste se a UI
mudar.
