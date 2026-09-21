import { afterEach, describe, expect, it, vi } from "vitest";
import { InativacaoApiError, postAnalisar, postExecutar, type AnaliseInativacao } from "./inativacaoApi";

const arquivo = (nome: string) => new File(["x"], nome);

const ANALISE: AnaliseInativacao = {
  usuarios: [],
  resumo: { executaveis: 0, estruturasExcluidas: 0, estruturasCompactadas: 0, estruturasOrfas: 0, duplicados: [] },
  impressaoDigital: "abc",
};

class FakeXhr {
  static respostas: { status: number; corpo: unknown }[] = [];
  static enviados: FormData[] = [];
  upload = { addEventListener: () => {} };
  responseType = "";
  status = 0;
  response: Blob | null = null;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  open() {}
  send(formData: FormData) {
    FakeXhr.enviados.push(formData);
    const proxima = FakeXhr.respostas.shift();
    queueMicrotask(() => {
      this.status = proxima?.status ?? 500;
      const corpo = proxima?.corpo ?? "";
      this.response = new Blob([typeof corpo === "string" ? corpo : JSON.stringify(corpo)]);
      this.onload?.();
    });
  }
}

afterEach(() => {
  vi.unstubAllGlobals();
  FakeXhr.respostas = [];
  FakeXhr.enviados = [];
});

describe("postAnalisar", () => {
  it("envia as bases, a lista e as escolhas, e devolve a análise", async () => {
    const fetchMock = vi.fn().mockResolvedValue({ ok: true, status: 200, json: async () => ANALISE });
    vi.stubGlobal("fetch", fetchMock);

    const out = await postAnalisar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["123"], ["456"]);

    expect(out).toEqual(ANALISE);
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe("/api/inativacao/analisar");
    const fd = init.body as FormData;
    expect(fd.get("itens")).toBe(JSON.stringify(["123"]));
    expect(fd.get("selecionados")).toBe(JSON.stringify(["456"]));
    expect((fd.get("cadastro") as File).name).toBe("c.xlsx");
    expect((fd.get("estruturas") as File).name).toBe("e.xlsx");
  });

  it("transforma a recusa do servidor em InativacaoApiError com o code", async () => {
    const recusa = { ok: false, status: 400, json: async () => ({ error: "Envie a base de estruturas.", code: "BASE_AUSENTE" }) };
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(recusa));

    const erro = await postAnalisar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["1"], []).catch((e: unknown) => e);

    expect(erro).toBeInstanceOf(InativacaoApiError);
    expect(erro).toMatchObject({ code: "BASE_AUSENTE", message: "Envie a base de estruturas." });
  });

  it("sem corpo JSON usa ERRO_INTERNO e o status na mensagem", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: false, status: 502, json: async () => Promise.reject(new Error("html")) }));
    const erro = await postAnalisar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["1"], []).catch((e: unknown) => e);
    expect(erro).toMatchObject({ code: "ERRO_INTERNO", message: "Falha na requisição (502)" });
  });
});

describe("postExecutar", () => {
  it("envia cpfs, impressão digital e a confirmação das órfãs, e devolve o ZIP", async () => {
    vi.stubGlobal("XMLHttpRequest", FakeXhr);
    FakeXhr.respostas.push({ status: 200, corpo: "conteudo-do-zip" });

    const blob = await postExecutar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["111", "222"], "digital", true);

    expect(blob.size).toBeGreaterThan(0);
    const fd = FakeXhr.enviados[0];
    expect(fd.get("cpfs")).toBe(JSON.stringify(["111", "222"]));
    expect(fd.get("impressaoDigital")).toBe("digital");
    expect(fd.get("ignore_orphan_warning")).toBe("true");
  });

  it("repassa o code do servidor (ex.: análise divergente)", async () => {
    vi.stubGlobal("XMLHttpRequest", FakeXhr);
    FakeXhr.respostas.push({ status: 409, corpo: { error: "A análise mudou.", code: "ANALISE_DIVERGENTE" } });

    const erro = await postExecutar(arquivo("c.xlsx"), arquivo("e.xlsx"), ["1"], "d", false).catch((e: unknown) => e);

    expect(erro).toBeInstanceOf(InativacaoApiError);
    expect(erro).toMatchObject({ code: "ANALISE_DIVERGENTE", message: "A análise mudou." });
  });
});
