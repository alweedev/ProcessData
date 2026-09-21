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
