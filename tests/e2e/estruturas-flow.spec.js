// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const APPROVER = validCpf(1);
const OTHER = validCpf(3);
const NEW_APPROVER = validCpf(4);

function usersRows() {
  return [
    { CPF: APPROVER, Status: "ATIVO", NomeCompleto: "Aprovador Um" },
    { CPF: OTHER, Status: "ATIVO", NomeCompleto: "Outro" },
    { CPF: NEW_APPROVER, Status: "ATIVO", NomeCompleto: "Aprovador Novo" },
  ];
}

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#estruturas-tab").click();
  await expect(page.locator("#estruturas")).toBeVisible();
});

test("verifica CPF, mostra a tabela e remove de todas as estruturas", async ({ page }) => {
  const base = [
    { AprovacaoId: "A1", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: OTHER },
  ];
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.locator("#aprovacao_preview_btn").click();

  await expect(page.locator("#aprovacao_approver_name")).toContainText("Aprovador Um");
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#aprovacao_remove_all_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("base_aprovacao_atualizada.xlsx");
});

test("avisa quando uma estrutura ficará sem aprovador e permite continuar", async ({ page }) => {
  // Sem LoginAprovador_2: remover o único aprovador esvazia a estrutura.
  const base = [{ AprovacaoId: "A2", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: "" }];
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_approver_name")).toContainText("Aprovador Um");

  await page.locator("#aprovacao_remove_all_btn").click();
  const confirmBtn = page.locator('[data-testid="aprovacao-confirm-continuar"]');
  await expect(confirmBtn).toBeVisible();

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    confirmBtn.click(),
  ]);
  expect(download.suggestedFilename()).toBe("base_aprovacao_atualizada.xlsx");
});

test("seleção por linha: remover selecionadas fica desabilitado sem seleção", async ({ page }) => {
  const base = [
    { AprovacaoId: "A3", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: OTHER },
  ];
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);

  // Todas vêm pré-selecionadas após o preview.
  await expect(page.locator("#aprovacao_remove_selected_btn")).toBeEnabled();

  await page.locator("#aprovacao_check_all").uncheck();
  await expect(page.locator("#aprovacao_remove_selected_btn")).toBeDisabled();
});

test("modo substituir: verifica os 2 CPFs, mostra tabela e substitui em todas as estruturas", async ({ page }) => {
  const base = [
    { AprovacaoId: "A4", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: OTHER },
  ];
  await page.locator("#aprovacao_mode_substituir").click();
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.fill("#aprovacao_new_cpf", NEW_APPROVER);
  await page.locator("#aprovacao_preview_btn").click();

  await expect(page.locator("#aprovacao_approver_name")).toContainText("Aprovador Um");
  await expect(page.locator("#aprovacao_new_approver_name")).toContainText("Aprovador Novo");
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#aprovacao_substitute_all_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("base_aprovacao_atualizada.xlsx");
});

test("modo substituir: avisa quando o novo aprovador já está na estrutura e permite continuar", async ({ page }) => {
  const base = [
    { AprovacaoId: "A5", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: NEW_APPROVER },
  ];
  await page.locator("#aprovacao_mode_substituir").click();
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.fill("#aprovacao_new_cpf", NEW_APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_approver_name")).toContainText("Aprovador Um");

  await page.locator("#aprovacao_substitute_all_btn").click();
  const confirmBtn = page.locator('[data-testid="aprovacao-confirm-continuar"]');
  await expect(confirmBtn).toBeVisible();

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    confirmBtn.click(),
  ]);
  expect(download.suggestedFilename()).toBe("base_aprovacao_atualizada.xlsx");
});
