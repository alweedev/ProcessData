// @ts-check
import { test, expect } from "@playwright/test";

// O log de auditoria do servidor é compartilhado entre specs (webServer
// reutilizado), então já existem outras linhas na tabela antes deste teste
// rodar. Usamos textos únicos e locators com hasText para mirar só na linha
// que este teste criou, em vez de assumir uma lista vazia.

/** Semeia uma entrada local (mesmo formato de historyStore.ts::addLocalEntry)
 * direto no localStorage e recarrega, pra HistoricoTab já nascer com o dado
 * -- evita depender de qualquer hook global exposto só para teste. */
async function seedHistoryEntry(page, text) {
  await page.evaluate((t) => {
    const KEY = "history_v2";
    const arr = JSON.parse(localStorage.getItem(KEY) || "[]");
    arr.push({ ts: Date.now(), text: t, source: "local" });
    localStorage.setItem(KEY, JSON.stringify(arr));
  }, text);
  await page.reload();
  await page.locator("#historico-tab").click();
  await expect(page.locator("#historico")).toBeVisible();
}

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#historico-tab").click();
  await expect(page.locator("#historico")).toBeVisible();
});

test("item aparece após addToHistory e busca filtra a lista", async ({ page }) => {
  const marker = `evento_teste_${Date.now()}`;
  await seedHistoryEntry(page, marker);

  const row = page.locator("#historico_tbody tr", { hasText: marker });
  await expect(row).toBeVisible();

  await page.fill("#historico_search", "não-existe-nada-com-esse-termo");
  await expect(page.locator("#historico_tbody tr")).toHaveCount(0);
  await expect(page.locator("#historico_summary")).toContainText("Itens visíveis: 0");

  await page.fill("#historico_search", marker);
  await expect(row).toBeVisible();
  await expect(page.locator("#historico_tbody tr")).toHaveCount(1);
});

test("exportar CSV e JSON disparam download", async ({ page }) => {
  const marker = `evento_export_${Date.now()}`;
  await seedHistoryEntry(page, marker);
  await expect(page.locator("#historico_tbody tr", { hasText: marker })).toBeVisible();

  const [csvDownload] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#historico_export_csv").click(),
  ]);
  expect(csvDownload.suggestedFilename()).toMatch(/^historico_\d+\.csv$/);

  const [jsonDownload] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#historico_export_json").click(),
  ]);
  expect(jsonDownload.suggestedFilename()).toMatch(/^historico_\d+\.json$/);
});

test("limpar esvazia a lista e chama DELETE /api/history", async ({ page }) => {
  const marker = `evento_para_limpar_${Date.now()}`;
  await seedHistoryEntry(page, marker);
  await expect(page.locator("#historico_tbody tr", { hasText: marker })).toBeVisible();

  const [request] = await Promise.all([
    page.waitForRequest((req) => req.url().includes("/api/history") && req.method() === "DELETE"),
    page.locator("#clearHistoryBtn").click(),
  ]);
  expect(request.method()).toBe("DELETE");

  await expect(page.locator("#historico_tbody tr")).toHaveCount(0);
  await expect(page.locator("#historico_summary")).toContainText("local: 0");
});
