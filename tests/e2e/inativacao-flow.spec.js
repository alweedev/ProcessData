// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const base = () => [
  { CPF: validCpf(1), NomeCompleto: "Maria Silva", Email: "maria@x.com", Status: "ATIVO" },
  { CPF: validCpf(2), NomeCompleto: "Joao Pereira", Email: "joao@x.com", Status: "ATIVO" },
];

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#inativacao-tab").click();
  await expect(page.locator("#inativacao")).toBeVisible(); // painel da aba
  await expect(page.locator("#inativacao_btn")).toBeVisible();
});

test("buscar e gerar saida_inativacao.xlsx; limpa a base no sucesso", async ({ page }) => {
  await page.setInputFiles("#inativacao_base", xlsxFile("base.xlsx", base()));
  await page.fill("#lista_text", validCpf(1));

  await page.locator("#inativacao_btn").click(); // buscar
  await expect(page.locator("#results_body tr")).not.toHaveCount(0);

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#inativacao_generate_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("saida_inativacao.xlsx");

  await expect
    .poll(async () => page.locator("#inativacao_base").evaluate((el) => el.files.length))
    .toBe(0);
});

test("drag-and-drop na area de upload da base", async ({ page }) => {
  // simula um drop preenchendo o input e disparando o change (equivalente ao handler)
  await page.setInputFiles("#inativacao_base", xlsxFile("base.xlsx", base()));
  await expect(page.locator("#inativacao_base_feedback")).toContainText("base.xlsx");
});
