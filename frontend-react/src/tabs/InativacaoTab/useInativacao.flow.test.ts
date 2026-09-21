import { act, renderHook } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import * as api from "../../lib/inativacaoApi";
import type { AnaliseInativacao } from "../../lib/inativacaoApi";
import { useInativacao } from "./useInativacao";

vi.mock("../../lib/inativacaoApi", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/inativacaoApi")>()),
  postAnalisar: vi.fn(),
  postExecutar: vi.fn(),
}));
vi.mock("../../lib/downloadFile", () => ({ triggerAnchorDownload: vi.fn() }));
vi.mock("../../runs/runsStore", () => ({ addRun: vi.fn() }));
vi.mock("../../toast/toastStore", () => ({ pushToast: vi.fn() }));

const CPF = "12345678909";
const ANALISE: AnaliseInativacao = {
  usuarios: [
    {
      cpf: CPF,
      cpfMascarado: "***.456.789-**",
      nome: "Ana Souza",
      email: "ana@x.com",
      situacao: "EXECUTAVEL",
      alerta: null,
      estruturasViajante: [],
      comoAprovador: [],
      candidatos: [],
    },
    {
      cpf: null,
      cpfMascarado: "",
      nome: "Sem Cpf",
      email: "",
      situacao: "SEM_CPF",
      alerta: "sem cpf",
      estruturasViajante: [],
      comoAprovador: [],
      candidatos: [],
    },
  ],
  resumo: { executaveis: 1, estruturasExcluidas: 0, estruturasCompactadas: 0, estruturasOrfas: 0, duplicados: [] },
  impressaoDigital: "digital-1",
};

const cadastro = new File(["c"], "cadastro.xlsx");
const estruturas = new File(["e"], "estruturas.xlsx");

function preparado() {
  const hook = renderHook(() => useInativacao());
  act(() => {
    hook.result.current.setCadastro(cadastro);
    hook.result.current.setEstruturas(estruturas);
    hook.result.current.setListText(`${CPF}\nAna Souza`);
  });
  return hook;
}

beforeEach(() => {
  vi.mocked(api.postAnalisar).mockReset();
  vi.mocked(api.postExecutar).mockReset();
  Object.assign(URL, { createObjectURL: vi.fn(() => "blob:zip") });
});

describe("useInativacao", () => {
  it("só habilita a análise com as duas bases e ao menos um item", () => {
    const { result } = renderHook(() => useInativacao());
    expect(result.current.podeAnalisar).toBe(false);
    act(() => {
      result.current.setCadastro(cadastro);
      result.current.setEstruturas(estruturas);
      result.current.setListText(CPF);
    });
    expect(result.current.podeAnalisar).toBe(true);
  });

  it("analisa enviando as bases e a lista já classificada", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();

    let ok = false;
    await act(async () => {
      ok = await result.current.analisar();
    });

    expect(ok).toBe(true);
    expect(api.postAnalisar).toHaveBeenCalledWith(cadastro, estruturas, [CPF, "Ana Souza"], []);
    expect(result.current.analise).toEqual(ANALISE);
  });

  it("guarda a falha da análise com a mensagem do servidor", async () => {
    vi.mocked(api.postAnalisar).mockRejectedValue(new api.InativacaoApiError("Envie a base.", "BASE_AUSENTE"));
    const { result } = preparado();

    let ok = true;
    await act(async () => {
      ok = await result.current.analisar();
    });

    expect(ok).toBe(false);
    expect(result.current.analise).toBeNull();
    expect(result.current.failure?.message).toBe("Envie a base.");
  });

  it("mudar a lista, uma base ou as escolhas de outra análise descarta a análise", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });
    expect(result.current.analise).not.toBeNull();

    act(() => result.current.setListText(CPF));
    expect(result.current.analise).toBeNull();
  });

  it("executa só os CPFs executáveis, baixa o ZIP e marca como concluído", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    vi.mocked(api.postExecutar).mockResolvedValue(new Blob(["zip"]));
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    let ok = false;
    await act(async () => {
      ok = await result.current.executar(true);
    });

    expect(ok).toBe(true);
    expect(api.postExecutar).toHaveBeenCalledWith(cadastro, estruturas, [CPF], "digital-1", true, expect.any(Function));
    expect(result.current.concluido).toBe(true);
  });

  it("análise divergente no servidor descarta a análise e guarda a falha", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    vi.mocked(api.postExecutar).mockRejectedValue(new api.InativacaoApiError("A análise mudou.", "ANALISE_DIVERGENTE"));
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    await act(async () => {
      await result.current.executar(false);
    });

    expect(result.current.analise).toBeNull();
    expect(result.current.concluido).toBe(false);
    expect(result.current.failure?.message).toBe("A análise mudou.");
  });

  it("escolher marca e desmarca homônimos", () => {
    const { result } = renderHook(() => useInativacao());
    act(() => result.current.escolher(CPF, true));
    expect(result.current.escolhidos).toEqual([CPF]);
    act(() => result.current.escolher(CPF, false));
    expect(result.current.escolhidos).toEqual([]);
  });

  it("reset volta ao estado inicial", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });
    act(() => result.current.reset());
    expect(result.current.analise).toBeNull();
    expect(result.current.cadastro).toBeNull();
    expect(result.current.listText).toBe("");
  });
});

function adiado<T>() {
  let resolve!: (valor: T) => void;
  let reject!: (erro: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

describe("useInativacao - respostas antigas e reentrada", () => {
  it("resposta da análise que chega depois de editar a lista é descartada", async () => {
    const pendente = adiado<AnaliseInativacao>();
    vi.mocked(api.postAnalisar).mockReturnValue(pendente.promise);
    const { result } = preparado();

    let emVoo: Promise<boolean> = Promise.resolve(true);
    act(() => {
      emVoo = result.current.analisar();
    });
    expect(result.current.analisando).toBe(true);

    act(() => result.current.setListText(CPF));
    expect(result.current.analisando).toBe(false);

    let ok = true;
    await act(async () => {
      pendente.resolve(ANALISE);
      ok = await emVoo;
    });

    expect(ok).toBe(false);
    expect(result.current.analise).toBeNull();
    expect(result.current.analisando).toBe(false);
    expect(result.current.failure).toBeNull();
  });

  it("erro da análise que chega depois de editar a lista também é descartado", async () => {
    const pendente = adiado<AnaliseInativacao>();
    vi.mocked(api.postAnalisar).mockReturnValue(pendente.promise);
    const { result } = preparado();

    let emVoo: Promise<boolean> = Promise.resolve(true);
    act(() => {
      emVoo = result.current.analisar();
    });
    act(() => result.current.setListText(CPF));
    await act(async () => {
      pendente.reject(new api.InativacaoApiError("Falhou.", "ERRO_INTERNO"));
      await emVoo;
    });

    expect(result.current.failure).toBeNull();
    expect(result.current.analisando).toBe(false);
  });

  it("resposta da análise que chega depois do reset é descartada", async () => {
    const pendente = adiado<AnaliseInativacao>();
    vi.mocked(api.postAnalisar).mockReturnValue(pendente.promise);
    const { result } = preparado();

    let emVoo: Promise<boolean> = Promise.resolve(true);
    act(() => {
      emVoo = result.current.analisar();
    });
    act(() => result.current.reset());

    await act(async () => {
      pendente.resolve(ANALISE);
      await emVoo;
    });

    expect(result.current.analise).toBeNull();
    expect(result.current.analisando).toBe(false);
  });

  it("uma análise antiga não apaga a flag de uma análise nova em andamento", async () => {
    const antiga = adiado<AnaliseInativacao>();
    const nova = adiado<AnaliseInativacao>();
    vi.mocked(api.postAnalisar).mockReturnValueOnce(antiga.promise).mockReturnValueOnce(nova.promise);
    const { result } = preparado();

    let p1: Promise<boolean> = Promise.resolve(true);
    let p2: Promise<boolean> = Promise.resolve(true);
    act(() => {
      p1 = result.current.analisar();
    });
    act(() => result.current.setListText(CPF));
    act(() => {
      p2 = result.current.analisar();
    });
    await act(async () => {
      antiga.resolve(ANALISE);
      await p1;
    });
    expect(result.current.analisando).toBe(true);
    expect(result.current.analise).toBeNull();

    await act(async () => {
      nova.resolve(ANALISE);
      await p2;
    });
    expect(result.current.analisando).toBe(false);
    expect(result.current.analise).toEqual(ANALISE);
  });

  it("analisar chamada duas vezes seguidas só envia uma requisição", async () => {
    const pendente = adiado<AnaliseInativacao>();
    vi.mocked(api.postAnalisar).mockReturnValue(pendente.promise);
    const { result } = preparado();

    let p1: Promise<boolean> = Promise.resolve(false);
    let segunda = true;
    await act(async () => {
      p1 = result.current.analisar();
      segunda = await result.current.analisar();
      pendente.resolve(ANALISE);
      await p1;
    });

    expect(segunda).toBe(false);
    expect(api.postAnalisar).toHaveBeenCalledTimes(1);
  });

  it("executar chamado duas vezes seguidas só dispara uma execução", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const pendente = adiado<Blob>();
    vi.mocked(api.postExecutar).mockReturnValue(pendente.promise);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    let p1: Promise<boolean> = Promise.resolve(false);
    let segunda = true;
    await act(async () => {
      p1 = result.current.executar(true);
      segunda = await result.current.executar(true);
      pendente.resolve(new Blob(["zip"]));
      await p1;
    });

    expect(segunda).toBe(false);
    expect(api.postExecutar).toHaveBeenCalledTimes(1);
    expect(result.current.concluido).toBe(true);
    expect(result.current.executando).toBe(false);
  });

  it("execução concluída depois de uma edição ainda baixa o ZIP e marca concluído", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const pendente = adiado<Blob>();
    vi.mocked(api.postExecutar).mockReturnValue(pendente.promise);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    let emVoo: Promise<boolean> = Promise.resolve(false);
    act(() => {
      emVoo = result.current.executar(true);
    });
    act(() => result.current.setListText(CPF));
    await act(async () => {
      pendente.resolve(new Blob(["zip"]));
      await emVoo;
    });

    expect(result.current.concluido).toBe(true);
    expect(result.current.executando).toBe(false);
  });

  it("falha de execução que chega depois de uma edição é ignorada", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const pendente = adiado<Blob>();
    vi.mocked(api.postExecutar).mockReturnValue(pendente.promise);
    const { result } = preparado();
    await act(async () => {
      await result.current.analisar();
    });

    let emVoo: Promise<boolean> = Promise.resolve(true);
    act(() => {
      emVoo = result.current.executar(false);
    });
    act(() => result.current.setListText(CPF));
    await act(async () => {
      pendente.reject(new api.InativacaoApiError("A análise mudou.", "ANALISE_DIVERGENTE"));
      await emVoo;
    });

    expect(result.current.failure).toBeNull();
    expect(result.current.executando).toBe(false);
  });

  it.each([
    ["setCadastro", (h: ReturnType<typeof useInativacao>) => h.setCadastro(new File(["outro"], "outro.xlsx"))],
    ["setEstruturas", (h: ReturnType<typeof useInativacao>) => h.setEstruturas(new File(["outro"], "outro.xlsx"))],
  ])("%s descarta a análise existente e as escolhas", async (_nome, mudar) => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    const { result } = preparado();
    act(() => result.current.escolher(CPF, true));
    await act(async () => {
      await result.current.analisar();
    });
    expect(result.current.analise).not.toBeNull();

    act(() => mudar(result.current));

    expect(result.current.analise).toBeNull();
    expect(result.current.escolhidos).toEqual([]);
    expect(result.current.failure).toBeNull();
  });
});
