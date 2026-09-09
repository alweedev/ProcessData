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
    command: "python -m backend.app",
    url: `${BASE_URL}/api/health`,
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
    env: {
      HOST: "127.0.0.1",
      PORT,
      UPLOAD_FOLDER: path.join(os.tmpdir(), "processdata_e2e_uploads"),
      HISTORY_LOG_FILE: path.join(os.tmpdir(), "processdata_e2e_uploads", "history.log.jsonl"),
    },
  },
});
