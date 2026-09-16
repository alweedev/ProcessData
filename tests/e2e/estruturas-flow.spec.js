// @ts-check
import { test, expect } from "@playwright/test";
import * as XLSX from "xlsx";
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

/** Lê as linhas do .xlsx baixado pelo browser. */
async function readDownloadRows(download) {
  const path = await download.path();
  const wb = XLSX.readFile(path);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function digitsOnly(value) {
  return String(value ?? "").replace(/\D/g, "");
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

test("toggle Remover/Substituir expõe aria-pressed corretamente", async ({ page }) => {
  const remover = page.locator("#aprovacao_mode_remover");
  const substituir = page.locator("#aprovacao_mode_substituir");
  await expect(remover).toHaveAttribute("aria-pressed", "true");
  await expect(substituir).toHaveAttribute("aria-pressed", "false");

  await substituir.click();
  await expect(substituir).toHaveAttribute("aria-pressed", "true");
  await expect(remover).toHaveAttribute("aria-pressed", "false");
});

test("modal de aviso tem focus trap: Tab não escapa pro conteúdo por trás", async ({ page }) => {
  const base = [{ AprovacaoId: "A2b", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: "" }];
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await page.locator("#aprovacao_remove_all_btn").click();

  const dialog = page.locator('[role="dialog"]');
  await expect(dialog).toBeVisible();

  // Foco inicial vai pro painel do diálogo; Shift+Tab a partir dele deve
  // "dar a volta" pro último elemento focável (o botão "Sim, continuar"),
  // não escapar pro restante da página.
  await page.keyboard.press("Shift+Tab");
  await expect(page.locator('[data-testid="aprovacao-confirm-continuar"]')).toBeFocused();

  // De lá, Tab avança normalmente pro botão "Fechar" (X) do diálogo -- ainda
  // dentro do modal, nunca pro input de CPF que fica atrás dele.
  await page.keyboard.press("Tab");
  await expect(page.locator("#aprovacao_cpf")).not.toBeFocused();
  await expect(dialog.locator(':focus')).toHaveCount(1);
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

test("modo remover: seleção por linha restringe a exportação às estruturas marcadas", async ({ page }) => {
  const base = [
    { AprovacaoId: "A9", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: OTHER },
    { AprovacaoId: "A10", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER, LoginAprovador_2: OTHER },
  ];
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);

  // Todas vêm pré-selecionadas após o preview; mantém só A9 marcada.
  await page.locator("#aprovacao_check_all").uncheck();
  await page.locator('.aprov-row-check[data-id="A9"]').check();

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#aprovacao_remove_selected_btn").click(),
  ]);
  const rows = await readDownloadRows(download);
  expect(rows).toHaveLength(1);
  expect(rows[0].AprovacaoId).toBe("A9");
  // APPROVER removido e compactado: OTHER sobe pra posição 1.
  expect(digitsOnly(rows[0].LoginAprovador_1)).toBe(OTHER);
  expect(String(rows[0].LoginAprovador_2 || "")).toBe("");
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

test("modo substituir: toggle 'segundo nível' controla se o LoginAprovador_SEGUNDO_NIVEL é substituído", async ({
  page,
}) => {
  const base = [
    {
      AprovacaoId: "A6",
      AprovacaoPor: "VIAJANTE",
      LoginAprovador_1: OTHER,
      LoginAprovador_SEGUNDO_NIVEL: APPROVER,
    },
  ];
  await page.locator("#aprovacao_mode_substituir").click();
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.fill("#aprovacao_new_cpf", NEW_APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);

  // Checkbox desligado (padrão): o segundo nível não é tocado.
  const [downloadOff] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#aprovacao_substitute_all_btn").click(),
  ]);
  const rowsOff = await readDownloadRows(downloadOff);
  expect(digitsOnly(rowsOff[0].LoginAprovador_SEGUNDO_NIVEL)).toBe(APPROVER);
  expect(String(rowsOff[0].Operacao || "")).toBe("");

  // Liga o toggle, refaz a verificação e exporta de novo.
  await page.locator("#aprovacao_replace_second_level").check();
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);
  const [downloadOn] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#aprovacao_substitute_all_btn").click(),
  ]);
  const rowsOn = await readDownloadRows(downloadOn);
  expect(digitsOnly(rowsOn[0].LoginAprovador_SEGUNDO_NIVEL)).toBe(NEW_APPROVER);
  expect(rowsOn[0].Operacao).toBe("UPDATE");
});

test("modo substituir: seleção por linha restringe a exportação às estruturas marcadas", async ({ page }) => {
  const base = [
    { AprovacaoId: "A7", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER },
    { AprovacaoId: "A8", AprovacaoPor: "VIAJANTE", LoginAprovador_1: APPROVER },
  ];
  await page.locator("#aprovacao_mode_substituir").click();
  await page.setInputFiles("#aprovacao_users_file", xlsxFile("users.xlsx", usersRows()));
  await page.setInputFiles("#aprovacao_base_file", xlsxFile("base.xlsx", base));
  await page.fill("#aprovacao_cpf", APPROVER);
  await page.fill("#aprovacao_new_cpf", NEW_APPROVER);
  await page.locator("#aprovacao_preview_btn").click();
  await expect(page.locator("#aprovacao_table_wrap tr")).not.toHaveCount(0);

  // Todas vêm pré-selecionadas após o preview; mantém só A7 marcada.
  await page.locator("#aprovacao_check_all").uncheck();
  await page.locator('.aprov-row-check[data-id="A7"]').check();
  await expect(page.locator("#aprovacao_substitute_selected_btn")).toBeEnabled();

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#aprovacao_substitute_selected_btn").click(),
  ]);
  const rows = await readDownloadRows(download);
  expect(rows).toHaveLength(1);
  expect(rows[0].AprovacaoId).toBe("A7");
  expect(digitsOnly(rows[0].LoginAprovador_1)).toBe(NEW_APPROVER);
});
