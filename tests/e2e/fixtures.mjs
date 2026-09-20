import * as XLSX from "xlsx";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Gera um .xlsx em memória a partir de uma lista de objetos (linhas). */
export function xlsxBuffer(rows) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dados");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

/** Payload pronto para page.setInputFiles(). */
export function xlsxFile(name, rows) {
  return { name, mimeType: XLSX_MIME, buffer: xlsxBuffer(rows) };
}

/** CPF de 11 dígitos com dígito verificador válido (determinístico por seed). */
export function validCpf(seed = 0) {
  const n = 100_000_000 + ((seed * 7_654_321) % 800_000_000);
  const base = String(n).padStart(9, "0");
  const digits = base.split("").map(Number);
  for (const pos of [9, 10]) {
    let sum = 0;
    for (let i = 0; i < pos; i++) sum += digits[i] * (pos + 1 - i);
    let dv = (sum * 10) % 11;
    if (dv === 10) dv = 0;
    digits[pos] = dv;
  }
  return digits.join("");
}

/** O Cadastro não tem valor padrão nem lembra a última escolha: tipo de login e
 *  fluxo precisam ser escolhidos a cada cadastro. */
export async function escolherConfigCadastro(page, login = "CPF", fluxo = "SELF") {
  await page.locator(`#cadastro_login_choice-${login}`).click();
  await page.locator(`#cadastro_fluxo-${fluxo}`).click();
}
