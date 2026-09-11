// Smoke visual: sobe o Flask, tira prints de cada tela em claro/escuro e em
// 1280 / 400 de largura. NÃO é gate — é pra revisão humana do re-skin.
// Fica FORA de tests/e2e/ pra o Playwright runner não coletar.
//
//   node tests/visual/smoke.mjs
//
// Saída: test-results-visual/<tela>-<tema>-<largura>.png (gitignored).
// (fora de test-results/ porque o Playwright limpa esse diretório a cada run.)
import { spawn } from "node:child_process";
import { mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const OUT = join(ROOT, "test-results-visual");
const PORT = process.env.SMOKE_PORT || "5099";
const BASE = `http://127.0.0.1:${PORT}`;
const PYTHON = process.env.PYTHON || join(ROOT, ".venv", "Scripts", "python.exe");

const VIEWS = ["home", "cadastro", "inativacao", "estruturas", "historico"];
const THEMES = ["light", "dark"];
const WIDTHS = [1280, 400];

async function waitForServer(url, timeoutMs = 30000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(url);
      if (res.ok) return;
    } catch {
      /* ainda subindo */
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  throw new Error(`servidor não respondeu em ${url}`);
}

async function main() {
  mkdirSync(OUT, { recursive: true });

  const server = spawn(PYTHON, ["-m", "backend.app"], {
    cwd: ROOT,
    env: { ...process.env, HOST: "127.0.0.1", PORT, PYTHONUTF8: "1" },
    stdio: "inherit",
  });

  let browser;
  try {
    await waitForServer(`${BASE}/api/health`);
    browser = await chromium.launch();

    for (const theme of THEMES) {
      for (const width of WIDTHS) {
        const context = await browser.newContext({ viewport: { width, height: 900 } });
        await context.addInitScript((t) => {
          try {
            localStorage.setItem("pd_theme", t);
          } catch {
            /* noop */
          }
        }, theme);
        const page = await context.newPage();
        for (const view of VIEWS) {
          const hash = view === "home" ? "#/" : `#/${view}`;
          await page.goto(`${BASE}/${hash}`, { waitUntil: "networkidle" });
          await page.waitForTimeout(250);
          const file = join(OUT, `${view}-${theme}-${width}.png`);
          await page.screenshot({ path: file, fullPage: true });
          console.log("  ✓", file.replace(ROOT + "\\", "").replace(ROOT + "/", ""));
        }
        await context.close();
      }
    }
  } finally {
    if (browser) await browser.close();
    server.kill();
  }
  console.log(`\n${VIEWS.length * THEMES.length * WIDTHS.length} prints em ${OUT}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
