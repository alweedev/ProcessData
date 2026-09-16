import { describe, expect, it } from "vitest";
import { validateCadastroFiles } from "./useCadastro";

function fakeFile(name: string, sizeBytes: number): File {
  const file = new File(["x"], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  Object.defineProperty(file, "size", { value: sizeBytes });
  return file;
}

describe("validateCadastroFiles", () => {
  it("aceita até 5 arquivos dentro do limite de tamanho", () => {
    const files = Array.from({ length: 5 }, (_, i) => fakeFile(`f${i}.xlsx`, 1024));
    expect(validateCadastroFiles(files)).toBeNull();
  });

  it("rejeita mais de 5 arquivos", () => {
    const files = Array.from({ length: 6 }, (_, i) => fakeFile(`f${i}.xlsx`, 1024));
    expect(validateCadastroFiles(files)).toMatch(/Máximo de 5 arquivos/);
  });

  it("rejeita um arquivo maior que 10 MB", () => {
    const bigFile = fakeFile("grande.xlsx", 11 * 1024 * 1024);
    expect(validateCadastroFiles([bigFile])).toMatch(/excede 10 MB/);
    expect(validateCadastroFiles([bigFile])).toContain("grande.xlsx");
  });

  it("lista vazia é válida (a UI trata 'nenhum arquivo' separadamente)", () => {
    expect(validateCadastroFiles([])).toBeNull();
  });
});
