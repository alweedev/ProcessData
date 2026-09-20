// @ts-check
import { test, expect } from "@playwright/test";
import { escolherConfigCadastro, xlsxFile, validCpf } from "./fixtures.mjs";

const cadastroRows = (n) =>
  Array.from({ length: n }, (_, i) => ({
    CPF: validCpf(i + 1),
    "NOME COMPLETO": `Pessoa ${i + 1}`,
    EMAIL: `p${i + 1}@x.com`,
    EMPRESA: "Empresa A",
    "Centro de custo": "CC1",
    "SOLICITANTE? (S/N)": "S",
  }));

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#cadastro-tab").click();
  await expect(page.locator("#cadastro")).toBeVisible(); // painel da aba
  await expect(page.locator("#cadastro_btn")).toBeVisible();
});

test("gera saida_cadastro.xlsx e limpa a selecao no sucesso", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", cadastroRows(1)));
  await escolherConfigCadastro(page);

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#cadastro_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("saida_cadastro.xlsx");

  // A mensagem final descreve o estado real: o cadastro JÁ está concluído (não
  // "o download começou") e cita o mesmo arquivo que o navegador recebeu.
  const status = page.locator("#cadastro_status");
  await expect(status).toContainText("Cadastro concluído");
  await expect(status).toContainText(download.suggestedFilename());
  await expect(status).not.toContainText("começou");

  // selecao limpa apos sucesso
  await expect
    .poll(async () => page.locator("#cadastro_files").evaluate((el) => el.files.length))
    .toBe(0);
});

test("sem planilha, validar e gerar ficam bloqueados", async ({ page }) => {
  await expect(page.locator("#cadastro_validate_btn")).toBeDisabled();
  await expect(page.locator("#cadastro_btn")).toBeDisabled();
});

test("com planilha mas sem escolher login e fluxo, continua bloqueado e diz o que falta", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", cadastroRows(1)));

  await expect(page.locator("#cadastro_btn")).toBeDisabled();
  await expect(page.locator("#cadastro_validate_btn")).toBeDisabled();
  await expect(page.locator("#cadastro_hint")).toContainText("tipo de login");
  await expect(page.locator("#cadastro_hint")).toContainText("fluxo");

  await page.locator("#cadastro_login_choice-CPF").click();
  await expect(page.locator("#cadastro_btn")).toBeDisabled();
  await expect(page.locator("#cadastro_hint")).not.toContainText("tipo de login");
  await expect(page.locator("#cadastro_hint")).toContainText("fluxo");

  await page.locator("#cadastro_fluxo-FRONT").click();
  await expect(page.locator("#cadastro_btn")).toBeEnabled();
  await expect(page.locator("#cadastro_validate_btn")).toBeEnabled();
});

test("as opções começam sem nenhuma escolha", async ({ page }) => {
  for (const id of ["CPF", "EMAIL"]) {
    await expect(page.locator(`#cadastro_login_choice-${id}`)).toHaveAttribute("aria-checked", "false");
  }
  for (const id of ["SELF", "FRONT"]) {
    await expect(page.locator(`#cadastro_fluxo-${id}`)).toHaveAttribute("aria-checked", "false");
  }
});

test("as escolhas não ficam salvas: ao fechar e reabrir é preciso escolher de novo", async ({ page }) => {
  await escolherConfigCadastro(page, "EMAIL", "FRONT");
  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "true");

  await page.reload();
  await page.locator("#cadastro-tab").click();

  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "false");
  await expect(page.locator("#cadastro_fluxo-FRONT")).toHaveAttribute("aria-checked", "false");
});

test("depois de gerar, as escolhas voltam a zero para o próximo cadastro", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", cadastroRows(1)));
  await escolherConfigCadastro(page, "EMAIL", "FRONT");

  await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  await expect(page.locator("#cadastro_status")).toContainText("Cadastro concluído");

  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "false");
  await expect(page.locator("#cadastro_fluxo-FRONT")).toHaveAttribute("aria-checked", "false");
});

test("depois de validar, o relatório aparece dentro da área visível da tela", async ({ page }) => {
  // Viewport baixa: o relatório nasce abaixo dos botões, fora da tela.
  await page.setViewportSize({ width: 1280, height: 560 });
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", cadastroRows(3)));
  await escolherConfigCadastro(page);

  await page.locator("#cadastro_validate_btn").click();

  await expect(page.getByText("A validação é informativa")).toBeInViewport();
});

test("bloqueia mais de 5 arquivos no cliente", async ({ page }) => {
  const files = cadastroRows(6).map((r, i) => xlsxFile(`c${i}.xlsx`, [r]));
  await page.setInputFiles("#cadastro_files", files);
  // guarda do cliente zera a selecao e mostra toast
  await expect
    .poll(async () => page.locator("#cadastro_files").evaluate((el) => el.files.length))
    .toBe(0);
});
