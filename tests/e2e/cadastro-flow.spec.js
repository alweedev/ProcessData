// @ts-check
import { test, expect } from "@playwright/test";
import { avancarAteGerar, xlsxFile, validCpf } from "./fixtures.mjs";

const cadastroRows = (n) =>
  Array.from({ length: n }, (_, i) => ({
    CPF: validCpf(i + 1),
    "NOME COMPLETO": `Pessoa ${i + 1}`,
    EMAIL: `p${i + 1}@x.com`,
    EMPRESA: "Empresa A",
    "Centro de custo": "CC1",
    "SOLICITANTE? (S/N)": "S",
  }));

const planilha = (n = 1) => xlsxFile("cadastro.xlsx", cadastroRows(n));

/** Botões da linha do tempo (só os pontos ao alcance viram botões). */
const pontosClicaveis = (page) => page.getByRole("list", { name: "Etapas do cadastro" }).getByRole("button");
const centroX = async (locator) => {
  const box = await locator.boundingBox();
  return box.x + box.width / 2;
};

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#cadastro-tab").click();
  await expect(page.locator("#cadastro")).toBeVisible(); // painel da aba
  await expect(page.locator("#cadastro_files")).toBeAttached(); // 1ª etapa: enviar fichas
});

test("gera saida_cadastro.xlsx e limpa a selecao no sucesso", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page);

  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  expect(download.suggestedFilename()).toBe("saida_cadastro.xlsx");

  // A mensagem final descreve o estado real: o cadastro JÁ está concluído (não
  // "o download começou") e cita o mesmo arquivo que o navegador recebeu.
  const status = page.locator("#cadastro_status");
  await expect(status).toContainText("Cadastro concluído");
  await expect(status).toContainText(download.suggestedFilename());
  await expect(status).not.toContainText("começou");
});

test("começa nas fichas e só avança quando há planilha", async ({ page }) => {
  await expect(page.locator("#cadastro_next_btn")).toBeDisabled();
  await expect(page.locator("#cadastro_back_btn")).toHaveCount(0); // não há etapa anterior

  await page.setInputFiles("#cadastro_files", planilha());

  await expect(page.locator("#cadastro_next_btn")).toBeEnabled();
});

test("escolher uma opção avança sozinho para a próxima etapa", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();

  await expect(page.locator("#cadastro_login_choice")).toBeVisible();
  await expect(page.locator("#cadastro_fluxo")).toHaveCount(0);
  // Sem valor padrão: nada vem marcado.
  await expect(page.locator("#cadastro_login_choice-CPF")).toHaveAttribute("aria-checked", "false");

  await page.locator("#cadastro_login_choice-CPF").click();
  await expect(page.locator("#cadastro_fluxo")).toBeVisible();
  await expect(page.locator("#cadastro_fluxo-SELF")).toHaveAttribute("aria-checked", "false");

  await page.locator("#cadastro_fluxo-SELF").click();
  await expect(page.locator("#cadastro_btn")).toBeVisible();
});

test("não dá para pular etapas: só os pontos já alcançados são clicáveis", async ({ page }) => {
  await expect(pontosClicaveis(page)).toHaveCount(0);

  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click(); // agora em "Tipo de login"

  await expect(pontosClicaveis(page)).toHaveCount(1); // só "Fichas"
  await expect(pontosClicaveis(page).first()).toContainText("Fichas");
});

test("Voltar mantém o que já foi escolhido", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();
  await page.locator("#cadastro_login_choice-EMAIL").click(); // avança para o fluxo

  await page.locator("#cadastro_back_btn").click();

  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "true");
});

test("clicar num ponto concluído volta à etapa; ao reescolher, vai direto ao fim se o resto já está pronto", async ({
  page,
}) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "CPF", "SELF");

  await pontosClicaveis(page).filter({ hasText: "Tipo de login" }).click();
  await expect(page.locator("#cadastro_login_choice-CPF")).toHaveAttribute("aria-checked", "true");

  await page.locator("#cadastro_login_choice-EMAIL").click();

  // O fluxo já estava escolhido: não faz parar de novo nele.
  await expect(page.locator("#cadastro_btn")).toBeVisible();
});

test("o anel da linha do tempo pula para a etapa atual", async ({ page }) => {
  const anel = page.getByTestId("stepper-marker");
  await expect.poll(async () => Math.abs((await centroX(anel)) - (await centroX(page.getByTestId("stepper-dot-0"))))).toBeLessThan(1);
  await expect(anel.locator("span")).toHaveCSS("animation-name", "none"); // abrir a tela não anima

  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();

  await expect(anel.locator("span")).toHaveCSS("animation-name", "pd-hop"); // o pulo
  await expect.poll(async () => Math.abs((await centroX(anel)) - (await centroX(page.getByTestId("stepper-dot-1"))))).toBeLessThan(1);
});

test("as escolhas não ficam salvas: ao fechar e reabrir é preciso escolher de novo", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "EMAIL", "FRONT");

  await page.reload();
  await page.locator("#cadastro-tab").click();
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();

  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "false");
  await page.locator("#cadastro_login_choice-CPF").click();
  await expect(page.locator("#cadastro_fluxo-FRONT")).toHaveAttribute("aria-checked", "false");
});

test("depois de gerar, 'Novo cadastro' recomeça nas fichas com tudo vazio", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "EMAIL", "FRONT");
  await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  await expect(page.locator("#cadastro_status")).toContainText("Cadastro concluído");
  await expect(pontosClicaveis(page)).toHaveCount(0); // concluído: só "Novo cadastro" recomeça

  await page.locator("#cadastro_new_btn").click();

  await expect(page.locator("#cadastro_files")).toBeAttached();
  await expect(page.locator("#cadastro_next_btn")).toBeDisabled();
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();
  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "false");
});

test("depois de validar, o relatório aparece dentro da área visível da tela", async ({ page }) => {
  // Viewport baixa: o relatório nasce abaixo dos botões, fora da tela.
  await page.setViewportSize({ width: 1280, height: 560 });
  await page.setInputFiles("#cadastro_files", planilha(3));
  await avancarAteGerar(page);

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
