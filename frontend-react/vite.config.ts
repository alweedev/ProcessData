import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const pkg = JSON.parse(
  readFileSync(fileURLToPath(new URL("./package.json", import.meta.url)), "utf-8"),
) as { version: string };

// Build de biblioteca (não um app Vite-servido): o resultado é injetado como
// <script type="module">/<link> dentro de frontend/index.html, que o Flask já
// serve como está (frontend/static/react/** fica acessível via /static/react/*
// sem nenhuma mudança no backend). Nomes de saída fixos (app-react.js/.css);
// cache-busting é feito pelo `?v=N` nas tags do index.html, bumpado a cada
// release (mesmo esquema dos favicons). Ferramenta interna, base de usuários
// pequena — não vale a máquina de filename com hash + reescrita do HTML.
export default defineConfig({
  plugins: [react()],
  base: "/static/react/",
  server: {
    proxy: {
      "/api": {
        target: "http://127.0.0.1:5000",
        changeOrigin: true,
      },
    },
  },
  // Build de biblioteca não aplica o `define` padrão de process.env.NODE_ENV
  // que o modo "app" do Vite faz automaticamente para libs como React/ReactDOM
  // (elas checam isso em runtime) — sem isso o bundle lança
  // "ReferenceError: process is not defined" assim que carrega no browser.
  define: {
    "process.env.NODE_ENV": '"production"',
    __APP_VERSION__: JSON.stringify(pkg.version),
  },
  build: {
    outDir: "../frontend/static/react",
    emptyOutDir: true,
    cssCodeSplit: false,
    lib: {
      entry: fileURLToPath(new URL("src/main.tsx", import.meta.url)),
      formats: ["es"],
      fileName: () => "app-react.js",
      cssFileName: "app-react",
    },
  },
});
