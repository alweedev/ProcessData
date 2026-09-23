// @ts-check
import { test, expect } from "@playwright/test";
import { xlsxFile, validCpf } from "./fixtures.mjs";

const cadastro = () => [
  { CPF: validCpf(1), NomeCompleto: "Maria Silva", Email: "maria@x.com", Status: "ATIVO" },
  { CPF: validCpf(2), NomeCompleto: "Joao Pereira", Email: "joao@x.com", Status: "ATIVO" },
];

const estrutura = (id, cpfViajante, ...aprovadores) => {
  const linha = { AprovacaoId: id, AprovacaoPor: "VIAJANTE", CPF: cpfViajante, NomeViajante: `Viajante ${id}` };
  aprovadores.forEach((cpf, i) => {
    linha[`LoginAprovador_${i + 1}`] = cpf;
  });
  return linha;
};

/** S1: estrutura direta de Maria. S2: Maria aprova junto com Joao (compactação). */
const estruturas = () => [estrutura("S1", validCpf(1), validCpf(2)), estrutura("S2", validCpf(7), validCpf(2), validCpf(1))];

/** S3: Maria é a única aprovadora (estrutura órfã). */
const estruturasComOrfa = () => [estrutura("S3", validCpf(7), validCpf(1))];

async function analisar(page, bases, lista) {
  await page.setInputFiles("#inativacao_cadastro", xlsxFile("cadastro.xlsx", cadastro()));
  await page.setInputFiles("#inativacao_estruturas", xlsxFile("estruturas.xlsx", bases));
  await page.fill("#lista_text", lista);
  await page.locator("#inativacao_btn").click();
  await expect(page.locator("#inativacao_impacto li")).not.toHaveCount(0);
}

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#inativacao-tab").click();
  await expect(page.locator("#inativacao")).toBeVisible(); // painel da aba
  await expect(page.locator("#inativacao_btn")).toBeDisabled(); // sem bases nem lista
});

test("analisar mostra o impacto e executar baixa o ZIP", async ({ page }) => {
  await analisar(page, estruturas(), validCpf(1));

  const card = page.locator("#inativacao_impacto li").first();
  await expect(card).toContainText("Será excluída: S1");
  await expect(card).toContainText("Compactação");

  await page.locator("#inativacao_next_btn").click();
  await expect(page.locator("#inativacao_execute_btn")).toBeDisabled(); // falta a confirmação
  await page.locator("#inativacao_confirm_impacto").check();

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#inativacao_execute_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("inativacao.zip");
  await expect(page.locator("#inativacao_status")).toContainText("Inativação concluída");
});

test("estrutura órfã exige uma segunda confirmação", async ({ page }) => {
  await analisar(page, estruturasComOrfa(), validCpf(1));
  await expect(page.locator("#inativacao_impacto li").first()).toContainText("Estrutura órfã");

  await page.locator("#inativacao_next_btn").click();
  await page.locator("#inativacao_confirm_orfas").check();
  await expect(page.locator("#inativacao_execute_btn")).toBeDisabled(); // ciência da órfã sozinha não basta
  await page.locator("#inativacao_confirm_impacto").check();
  await expect(page.locator("#inativacao_execute_btn")).toBeEnabled();
  await page.locator("#inativacao_confirm_orfas").uncheck();
  await expect(page.locator("#inativacao_execute_btn")).toBeDisabled(); // impacto sozinho não basta
  await page.locator("#inativacao_confirm_orfas").check();
  await expect(page.locator("#inativacao_execute_btn")).toBeEnabled();
});

test("usuário sem CPF no cadastro não pode ser inativado", async ({ page }) => {
  const semCpf = [{ CPF: "", NomeCompleto: "Sem Cpf Silva", Email: "s@x.com", Status: "ATIVO" }];
  await page.setInputFiles("#inativacao_cadastro", xlsxFile("cadastro.xlsx", semCpf));
  await page.setInputFiles("#inativacao_estruturas", xlsxFile("estruturas.xlsx", estruturas()));
  await page.fill("#lista_text", "Sem Cpf Silva");
  await page.locator("#inativacao_btn").click();

  await expect(page.locator("#inativacao_impacto li").first()).toContainText("não possui CPF registrado");
  await expect(page.locator("#inativacao_next_btn")).toBeDisabled();
});

test("editar a lista descarta a análise já feita", async ({ page }) => {
  const etapas = page.locator('ol[aria-label="Etapas da inativação"]');
  await analisar(page, estruturas(), validCpf(1));
  await expect(etapas).toContainText("1 a inativar");

  // só voltar de etapa não descarta a análise
  await page.locator("#inativacao_back_btn").click();
  await expect(etapas).toContainText("1 a inativar");

  // editar a lista descarta a análise
  await page.fill("#lista_text", validCpf(2));
  await expect(etapas).toContainText("Conferir o impacto");
  await expect(etapas).not.toContainText("1 a inativar");

  // e uma nova análise reflete a lista nova
  await page.locator("#inativacao_btn").click();
  const card = page.locator("#inativacao_impacto li").first();
  await expect(card).toContainText("Joao Pereira");
  await expect(card).not.toContainText("Maria Silva");
});
