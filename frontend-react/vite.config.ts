import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

// Build de biblioteca (não um app Vite-servido): o resultado é injetado como
// <script type="module">/<link> dentro de frontend/index.html, que o Flask já
// serve como está (frontend/static/react/** fica acessível via /static/react/*
// sem nenhuma mudança no backend). Nomes de saída fixos por enquanto (sem
// hash) para simplificar a migração incremental; hash volta na limpeza final.
export default defineConfig({
  plugins: [react()],
  base: "/static/react/",
  // Build de biblioteca não aplica o `define` padrão de process.env.NODE_ENV
  // que o modo "app" do Vite faz automaticamente para libs como React/ReactDOM
  // (elas checam isso em runtime) — sem isso o bundle lança
  // "ReferenceError: process is not defined" assim que carrega no browser.
  define: {
    "process.env.NODE_ENV": '"production"',
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
