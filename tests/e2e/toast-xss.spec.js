// @ts-check
import { test, expect } from "@playwright/test";
import { escolherConfigCadastro, xlsxFile, validCpf } from "./fixtures.mjs";

const XSS = '<img src=x onerror="window.__xss=1"> "><script>window.__xss=1</script>';

test("showToast renderiza mensagem maliciosa como texto, nunca como HTML", async ({ page }) => {
  // Regressão: a implementação legada inseria a mensagem via toast.innerHTML
  // sem escaping -- qualquer texto de erro vindo do servidor (ex.: um campo
  // refletido numa mensagem de validação) executaria como HTML. O toast
  // sempre renderiza via texto JSX (ToastViewport), que escapa por padrão.
  // Simulamos aqui o servidor devolvendo esse payload num erro real de
  // /api/process_cadastro, disparado pelo fluxo normal da UI (sem hook
  // global só-de-teste).
  let dialog = false;
  page.on("dialog", (d) => {
    dialog = true;
    d.dismiss().catch(() => {});
  });

  await page.route("**/api/process_cadastro", (route) =>
    route.fulfill({ status: 400, contentType: "application/json", body: JSON.stringify({ error: XSS }) }),
  );

  await page.goto("/");
  await page.locator("#cadastro-tab").click();
  await expect(page.locator("#cadastro_btn")).toBeVisible();
  await page.setInputFiles(
    "#cadastro_files",
    xlsxFile("cadastro.xlsx", [
      {
        CPF: validCpf(1),
        "NOME COMPLETO": "Pessoa Um",
        EMAIL: "pessoa1@x.com",
        EMPRESA: "Empresa A",
        "Centro de custo": "CC1",
        "SOLICITANTE? (S/N)": "S",
      },
    ]),
  );
  await escolherConfigCadastro(page);
  await page.locator("#cadastro_btn").click();

  // Não usa .first(): o ApiStatusBadge dispara um toast "API Online" ao
  // montar, então localiza especificamente o toast injetado por este teste.
  const toastBody = page.locator("#toastContainer .toast-body", { hasText: "onerror" });
  await expect(toastBody).toContainText("onerror");
  expect(await toastBody.evaluate((el) => el.querySelectorAll("img,script").length)).toBe(0);
  expect(dialog).toBe(false);
  expect(await page.evaluate(() => window.__xss)).toBeFalsy();
});
