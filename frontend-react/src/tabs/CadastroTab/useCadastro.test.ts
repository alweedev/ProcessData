import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import * as api from "../../lib/api";
import * as toasts from "../../toast/toastStore";
import { isSpreadsheet, useCadastro, validateCadastroFiles } from "./useCadastro";

vi.mock("../../lib/api", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/api")>()), // mantém ApiRejection real
  postFormForBlob: vi.fn(),
  postAnalysisSummary: vi.fn(),
}));
vi.mock("../../toast/toastStore", () => ({ pushToast: vi.fn() }));

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

  it("rejeita arquivo que não é planilha, apontando o nome", () => {
    expect(validateCadastroFiles([fakeFile("relatorio.pdf", 1024)])).toBe(
      '"relatorio.pdf" não é uma planilha .xlsx ou .xls.',
    );
  });
});

describe("isSpreadsheet", () => {
  it.each([
    ["fichas.xlsx", true],
    ["fichas.xls", true],
    ["FICHAS.XLSX", true], // extensão em maiúsculas
    ["fichas.xlsx.pdf", false], // só a última extensão conta
    ["fichas.csv", false],
    ["xlsx", false],
  ])("%s -> %s", (nome, esperado) => {
    expect(isSpreadsheet(fakeFile(nome, 1))).toBe(esperado);
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
    expect(hook.result.current.failure).toEqual({ message: "Coluna obrigatória ausente: CPF", fileErrors: undefined });
    return hook;
  }

  it("escolher novos arquivos limpa a mensagem de erro", async () => {
    const hook = await hookAfterFailedSubmit();
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("corrigido.xlsx", 2048)));
    });
    expect(hook.result.current.failure).toBeNull();
  });

  it("remover um arquivo limpa a mensagem de erro", async () => {
    const hook = await hookAfterFailedSubmit();
    act(() => {
      hook.result.current.removeFile(0);
    });
    expect(hook.result.current.failure).toBeNull();
  });

  it("uma seleção rejeitada (arquivos demais) também limpa a mensagem de erro", async () => {
    const hook = await hookAfterFailedSubmit();
    const seis = Array.from({ length: 6 }, (_, i) => fakeFile(`f${i}.xlsx`, 1024));
    let aceito = true;
    act(() => {
      aceito = hook.result.current.pickFiles(fileList(...seis));
    });
    expect(aceito).toBe(false);
    expect(hook.result.current.failure).toBeNull();
  });
});

describe("useCadastro — seleção de arquivos (arrastar e soltar não passa pelo `accept`)", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const nomes = (hook: { result: { current: ReturnType<typeof useCadastro> } }) =>
    hook.result.current.files.map((f) => f.name);

  it("descarta o que não é planilha, avisa pelo nome e segue com o resto", () => {
    const hook = renderHook(() => useCadastro());
    let aceito = false;
    act(() => {
      aceito = hook.result.current.pickFiles(
        fileList(fakeFile("a.xlsx", 1), fakeFile("notas.pdf", 1), fakeFile("b.xls", 1)),
      );
    });

    expect(aceito).toBe(true);
    expect(nomes(hook)).toEqual(["a.xlsx", "b.xls"]);
    expect(toasts.pushToast).toHaveBeenCalledExactlyOnceWith(
      '"notas.pdf" foi ignorado — só planilhas .xlsx ou .xls são aceitas.',
      "info",
    );
  });

  it("com vários ignorados cita o primeiro e a quantidade dos demais", () => {
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(
        fileList(fakeFile("a.xlsx", 1), fakeFile("x.pdf", 1), fakeFile("y.png", 1), fakeFile("z.txt", 1)),
      );
    });

    expect(toasts.pushToast).toHaveBeenCalledWith(
      '"x.pdf" e mais 2 arquivos foram ignorados — só planilhas .xlsx ou .xls são aceitas.',
      "info",
    );
  });

  it("sem nenhuma planilha a seleção é recusada e o que já estava escolhido continua", () => {
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1)));
    });
    let aceito = true;
    act(() => {
      aceito = hook.result.current.pickFiles(fileList(fakeFile("foto.png", 1)));
    });

    expect(aceito).toBe(false);
    expect(nomes(hook)).toEqual(["a.xlsx"]);
    expect(toasts.pushToast).toHaveBeenLastCalledWith(
      '"foto.png" não é uma planilha — só .xlsx ou .xls são aceitos.',
      "danger",
    );
  });

  it("arquivos demais também recusam a seleção sem apagar a anterior", () => {
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1)));
    });
    const seis = Array.from({ length: 6 }, (_, i) => fakeFile(`f${i}.xlsx`, 1024));
    let aceito = true;
    act(() => {
      aceito = hook.result.current.pickFiles(fileList(...seis));
    });

    expect(aceito).toBe(false);
    expect(nomes(hook)).toEqual(["a.xlsx"]);
  });

  it("uma seleção aceita descarta o relatório de validação da anterior", async () => {
    vi.mocked(api.postAnalysisSummary).mockResolvedValueOnce({ report: { total_rows: 1 } } as never);
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1)));
    });
    await act(async () => {
      await hook.result.current.validate("CPF", "SELF");
    });
    expect(hook.result.current.validation.status).toBe("done");

    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("b.xlsx", 1)));
    });

    expect(hook.result.current.validation.status).toBe("idle");
  });
});

describe("useCadastro — conferência de nomes", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
    vi.unstubAllGlobals();
  });

  const item = {
    key: "1:2",
    label: "Linha 2",
    nome_completo: "XANTO ZEKO QUIM BRAG",
    nome: "XANTO",
    sobrenome: "ZEKO QUIM BRAG",
    confianca: "baixa" as const,
    motivos: ["a divisão entre nome e sobrenome é ambígua"],
    estouro: { nome: false, sobrenome: false },
    sugestao: null,
  };

  it("a validação guarda a lista de nomes para conferir, e recomeçar a validação a descarta", async () => {
    vi.mocked(api.postAnalysisSummary).mockResolvedValueOnce({
      report: { total_rows: 1 },
      preview: [],
      name_review: [item],
    } as never);
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });

    await act(async () => {
      await hook.result.current.validate("CPF", "SELF");
    });
    expect(hook.result.current.validation.nameReview).toEqual([item]);

    act(() => {
      hook.result.current.invalidateValidation();
    });
    expect(hook.result.current.validation.nameReview).toEqual([]);
  });

  it("validação de servidor antigo (sem name_review) não quebra: lista vazia", async () => {
    vi.mocked(api.postAnalysisSummary).mockResolvedValueOnce({ report: { total_rows: 1 }, preview: [] } as never);
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });

    await act(async () => {
      await hook.result.current.validate("CPF", "SELF");
    });

    expect(hook.result.current.validation.nameReview).toEqual([]);
  });

  async function gerarCom(nameOverrides?: Parameters<ReturnType<typeof useCadastro>["submit"]>[3]) {
    vi.stubGlobal("URL", { ...URL, createObjectURL: () => "blob:teste", revokeObjectURL: () => {} });
    vi.mocked(api.postFormForBlob).mockResolvedValueOnce(new Blob(["x"]));
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });
    await act(async () => {
      await hook.result.current.submit("CPF", "SELF", undefined, nameOverrides);
    });
    return vi.mocked(api.postFormForBlob).mock.calls[0][1] as FormData;
  }

  it("a geração leva junto as decisões da conferência (JSON em name_overrides)", async () => {
    const overrides = {
      "1:2": { nome_completo: item.nome_completo, nome: "XANTO ZEKO", sobrenome: "QUIM BRAG", editado: true },
    };

    const form = await gerarCom(overrides);

    expect(JSON.parse(String(form.get("name_overrides")))).toEqual(overrides);
    expect(form.get("login_choice")).toBe("CPF");
  });

  it("sem decisões, name_overrides nem vai no envio", async () => {
    expect((await gerarCom({})).has("name_overrides")).toBe(false);
  });
});

describe("useCadastro — erro da validação", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  async function validarComFalha(erro: Error) {
    vi.mocked(api.postAnalysisSummary).mockRejectedValueOnce(erro);
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });
    await act(async () => {
      await hook.result.current.validate("CPF", "SELF");
    });
    return hook.result.current.validation;
  }

  it("planilha recusada pelo servidor mostra o motivo, sem dizer que a geração não depende da validação", async () => {
    const validation = await validarComFalha(
      new api.ApiRejection(
        "Falha ao processar arquivo(s): a.xlsx: A planilha não tem linhas de dados abaixo do cabeçalho.",
        {
          "a.xlsx": "A planilha não tem linhas de dados abaixo do cabeçalho.",
        },
      ),
    );

    expect(validation.status).toBe("error");
    expect(validation.error).toBeNull(); // sem o texto genérico "geração não depende disso"
    expect(validation.failure).toEqual({
      message: "Falha ao processar arquivo(s): a.xlsx: A planilha não tem linhas de dados abaixo do cabeçalho.",
      fileErrors: { "a.xlsx": "A planilha não tem linhas de dados abaixo do cabeçalho." },
    });
    expect(toasts.pushToast).not.toHaveBeenCalled(); // o motivo já está na tela; sem toast genérico por cima
  });

  it("falha de rede continua dizendo que a geração não depende da validação", async () => {
    const validation = await validarComFalha(new TypeError("Failed to fetch"));

    expect(validation.error).toBe("Não foi possível validar agora (Failed to fetch) — a geração não depende disso.");
    expect(toasts.pushToast).toHaveBeenCalledWith("Não foi possível validar a planilha agora.", "info");
  });
});

describe("useCadastro — validação some quando deixa de valer", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const relatorio = { total: 1 } as unknown as api.QualityReport;

  async function hookComRelatorio() {
    vi.mocked(api.postAnalysisSummary).mockResolvedValueOnce({ report: relatorio } as never);
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });
    await act(async () => {
      await hook.result.current.validate("CPF", "SELF");
    });
    expect(hook.result.current.validation.status).toBe("done");
    return hook;
  }

  it("invalidar (login/fluxo mudou) descarta o relatório de validação", async () => {
    const hook = await hookComRelatorio();
    act(() => {
      hook.result.current.invalidateValidation();
    });
    expect(hook.result.current.validation.status).toBe("idle");
    expect(hook.result.current.validation.report).toBeNull();
  });

  it("invalidar no meio da validação impede que a resposta atrasada reapareça", async () => {
    let responder!: (v: never) => void;
    vi.mocked(api.postAnalysisSummary).mockReturnValueOnce(new Promise((resolve) => (responder = resolve as never)));
    const hook = renderHook(() => useCadastro());
    act(() => {
      hook.result.current.pickFiles(fileList(fakeFile("a.xlsx", 1024)));
    });
    let pendente!: Promise<void>;
    act(() => {
      pendente = hook.result.current.validate("CPF", "SELF");
    });
    expect(hook.result.current.validation.status).toBe("loading");

    act(() => {
      hook.result.current.invalidateValidation();
    });
    await act(async () => {
      responder({ report: relatorio } as never);
      await pendente;
    });

    expect(hook.result.current.validation.status).toBe("idle");
  });

  it("limpar as fichas também descarta o relatório", async () => {
    const hook = await hookComRelatorio();
    act(() => {
      hook.result.current.clear();
    });
    expect(hook.result.current.validation.status).toBe("idle");
    expect(hook.result.current.files).toHaveLength(0);
  });
});
