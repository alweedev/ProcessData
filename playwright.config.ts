import { defineConfig, devices } from "@playwright/test";
import os from "node:os";
import path from "node:path";

const PORT = process.env.E2E_PORT || "5001";
const BASE_URL = `http://127.0.0.1:${PORT}`;

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  fullyParallel: false,
  workers: 1,
  reporter: [["list"]],
  use: {
    baseURL: BASE_URL,
    trace: "on-first-retry",
  },
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"] } }],
  webServer: {
    // defina PYTHON para o interpretador do venv se "python" não estiver no PATH
    command: `${process.env.PYTHON || "python"} -m backend.app`,
    url: `${BASE_URL}/api/health`,
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
    env: {
      HOST: "127.0.0.1",
      PORT,
      UPLOAD_FOLDER: path.join(os.tmpdir(), "processdata_e2e_uploads"),
      HISTORY_LOG_FILE: path.join(os.tmpdir(), "processdata_e2e_uploads", "history.log.jsonl"),
      // O vocabulário de nomes aprende com as conferências: os testes não podem gravar no do usuário, e cada
      // execução começa de um vocabulário limpo (senão o que uma rodada ensinou mudaria a seguinte).
      NAME_VOCAB_FILE: path.join(os.tmpdir(), "processdata_e2e_uploads", `name_vocabulary_${Date.now()}.json`),
    },
  },
});
