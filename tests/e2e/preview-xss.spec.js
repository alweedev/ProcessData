// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const XSS = '<img src=x onerror="window.__xss=1"> "><script>window.__xss=1</script>';

test("nome malicioso da planilha aparece como texto na busca de inativacao", async ({ page }) => {
  let dialog = false;
  page.on("dialog", (d) => {
    dialog = true;
    d.dismiss().catch(() => {});
  });

  await page.goto("/");
  await page.locator("#inativacao-tab").click();

  const base = [
    { CPF: validCpf(1), NomeCompleto: XSS, Email: "a@x.com", Status: "ATIVO" },
  ];
  await page.setInputFiles("#inativacao_base", xlsxFile("base.xlsx", base));
  await page.fill("#lista_text", validCpf(1));
  await page.locator("#inativacao_btn").click();

  const cell = page.locator("#results_body tr td").first();
  await expect(cell).toContainText("onerror"); // renderizado como texto
  expect(await cell.evaluate((el) => el.querySelectorAll("img,script").length)).toBe(0);
  expect(dialog).toBe(false);
  expect(await page.evaluate(() => window.__xss)).toBeFalsy();
});
