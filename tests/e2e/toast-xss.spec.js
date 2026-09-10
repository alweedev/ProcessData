// @ts-check
import { test, expect } from "@playwright/test";

const XSS = '<img src=x onerror="window.__xss=1"> "><script>window.__xss=1</script>';

test("showToast renderiza mensagem maliciosa como texto, nunca como HTML", async ({ page }) => {
  // Regressão: a implementação legada inseria a mensagem via toast.innerHTML
  // sem escaping — qualquer texto de erro vindo do servidor (ex.: um campo
  // refletido numa mensagem de validação) executaria como HTML. A ponte
  // window.showToast (src/toast/legacyBridge.ts) agora renderiza via texto
  // JSX (ToastViewport), que escapa por padrão.
  let dialog = false;
  page.on("dialog", (d) => {
    dialog = true;
    d.dismiss().catch(() => {});
  });

  await page.goto("/");
  await page.evaluate((msg) => window.showToast(msg, "danger"), XSS);

  // Não usa .first(): o ApiStatusBadge dispara um toast "API Online" ao
  // montar, então localiza especificamente o toast injetado por este teste.
  const toastBody = page.locator("#toastContainer .toast-body", { hasText: "onerror" });
  await expect(toastBody).toContainText("onerror");
  expect(await toastBody.evaluate((el) => el.querySelectorAll("img,script").length)).toBe(0);
  expect(dialog).toBe(false);
  expect(await page.evaluate(() => window.__xss)).toBeFalsy();
});
