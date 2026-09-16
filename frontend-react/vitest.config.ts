import react from "@vitejs/plugin-react";
import { defineConfig } from "vitest/config";

// Config separada de vite.config.ts: aquele é um build de biblioteca
// (build.lib, sem dev server) e não serve pra rodar testes.
export default defineConfig({
  plugins: [react()],
  test: {
    environment: "jsdom",
    setupFiles: ["./src/setupTests.ts"],
    css: false,
  },
});
