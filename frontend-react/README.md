# ProcessData — frontend

Interface React do ProcessData (ferramentas internas do suporte Vermari:
cadastro em massa, inativação, estruturas de aprovação, histórico).

## Como funciona o build

`npm run build` roda `tsc -b && vite build` em **modo biblioteca**: a saída são
os nomes fixos `../frontend/static/react/app-react.{js,css}`, que
`../frontend/index.html` injeta como `<script type="module">` / `<link>`. O
Flask serve `frontend/` como está — `/static/react/*` fica acessível **sem
nenhuma mudança no backend**. Cache-busting é o `?v=N` nas tags do
`frontend/index.html` (bumpado a cada release, mesmo esquema dos favicons).

> O backend Flask **nunca** é modificado a partir daqui.

## Scripts

| Comando | O quê |
|---|---|
| `npm run build` (ou `npm run build:react` na raiz) | build de produção para `../frontend/static/react/` |
| `npm run dev` (ou `npm run dev:react -- --port 5173` na raiz) | Vite dev server com HMR |
| `npm run lint` | oxlint |

## Dev com HMR

1. `npm run dev:react` (na raiz do repo) sobe o Vite em `localhost:5173`.
2. No console do navegador (na página servida pelo Flask):
   `localStorage.setItem('pd_dev_hmr', '1')` e recarregue.
   O `frontend/index.html` passa a carregar os módulos do dev server
   (`@vite/client` + Fast Refresh) em vez do build estático.
3. Para voltar ao build estático: `localStorage.removeItem('pd_dev_hmr')`.

`index.html` na raiz de `frontend-react/` só é usado pelo `vite dev`/`preview` —
não é o host de produção.

## Tema

Classe `.dark` no `<html>`, aplicada antes do 1º paint por um script anti-FOUC
no `<head>` do `frontend/index.html` (lê `localStorage.pd_theme`) e depois pelo
`ThemeToggle`. Os tokens de cor (`--pd-*` em `src/index.css`) trocam por essa
classe, então os componentes não precisam de variante `dark:`.
