// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const XSS = '<img src=x onerror="window.__xss=1"> "><script>window.__xss=1</script>';

test("nome malicioso da planilha aparece como texto na análise de inativacao", async ({ page }) => {
  let dialog = false;
  page.on("dialog", (d) => {
    dialog = true;
    d.dismiss().catch(() => {});
  });

  await page.goto("/");
  await page.locator("#inativacao-tab").click();

  const cadastro = [{ CPF: validCpf(1), NomeCompleto: XSS, Email: "a@x.com", Status: "ATIVO" }];
  const estruturas = [{ AprovacaoId: "S1", AprovacaoPor: "VIAJANTE", CPF: validCpf(1), LoginAprovador_1: validCpf(2) }];
  await page.setInputFiles("#inativacao_cadastro", xlsxFile("cadastro.xlsx", cadastro));
  await page.setInputFiles("#inativacao_estruturas", xlsxFile("estruturas.xlsx", estruturas));
  await page.fill("#lista_text", validCpf(1));
  await page.locator("#inativacao_btn").click();

  const card = page.locator("#inativacao_impacto li").first();
  await expect(card).toContainText("onerror"); // renderizado como texto
  expect(await card.evaluate((el) => el.querySelectorAll("img,script").length)).toBe(0);
  expect(dialog).toBe(false);
  expect(await page.evaluate(() => window.__xss)).toBeFalsy();
});
