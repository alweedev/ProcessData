import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import * as api from "../../lib/api";
import { useCadastro, validateCadastroFiles } from "./useCadastro";

vi.mock("../../lib/api", () => ({ postFormForBlob: vi.fn(), postAnalysisSummary: vi.fn() }));

function fakeFile(name: string, sizeBytes: number): File {
  const file = new File(["x"], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  Object.defineProperty(file, "size", { value: sizeBytes });
  return file;
}

function fileList(...files: File[]): FileList {
  return files as unknown as FileList;
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

describe("useCadastro — erro da tentativa anterior", () => {
  // Este projeto roda vitest sem `test.globals`, então o cleanup do Testing
  // Library não se registra sozinho (ver useHashRoute.test.ts).
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  async function hookAfterFailedSubmit() {
    vi.mocked(api.postFormForBlob).mockRejectedValueOnce(new Error("Coluna obrigatória ausente: CPF"));
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });
    await act(async () => {
      await hook.result.current.submit("CPF", "SELF");
    });
    expect(hook.result.current.debugMsg).toBe("Erro: Coluna obrigatória ausente: CPF");
    return hook;
  }

  it("escolher novos arquivos limpa a mensagem de erro", async () => {
    const hook = await hookAfterFailedSubmit();
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("corrigido.xlsx", 2048)));
    });
    expect(hook.result.current.debugMsg).toBe("");
  });

  it("remover um arquivo limpa a mensagem de erro", async () => {
    const hook = await hookAfterFailedSubmit();
    act(() => {
      hook.result.current.removeFile(0);
    });
    expect(hook.result.current.debugMsg).toBe("");
  });

  it("uma seleção rejeitada (arquivos demais) também limpa a mensagem de erro", async () => {
    const hook = await hookAfterFailedSubmit();
    const seis = Array.from({ length: 6 }, (_, i) => fakeFile(`f${i}.xlsx`, 1024));
    let aceito = true;
    act(() => {
      aceito = hook.result.current.pickFiles(fileList(...seis));
    });
    expect(aceito).toBe(false);
    expect(hook.result.current.debugMsg).toBe("");
  });
});
