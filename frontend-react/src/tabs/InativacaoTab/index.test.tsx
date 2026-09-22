import { cleanup, render, screen, waitFor } from "@testing-library/react";
import type { ReactNode } from "react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import * as api from "../../lib/inativacaoApi";
import type { AnaliseInativacao } from "../../lib/inativacaoApi";
import { InativacaoTab } from ".";

vi.mock("../../lib/inativacaoApi", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/inativacaoApi")>()),
  postAnalisar: vi.fn(),
  postExecutar: vi.fn(),
}));
// O FileDropzone real espelha o input com DataTransfer, que o jsdom não tem: aqui basta um <input type="file">.
vi.mock("../../components/FileDropzone", () => ({
  FileDropzone: ({ id, onFiles, children }: { id: string; onFiles: (f: FileList) => void; children?: ReactNode }) => (
    <div>
      <input id={id} type="file" onChange={(e) => e.target.files && onFiles(e.target.files)} />
      {children}
    </div>
  ),
}));
vi.mock("../../lib/downloadFile", () => ({ triggerAnchorDownload: vi.fn() }));
vi.mock("../../runs/runsStore", () => ({ addRun: vi.fn(), useRuns: () => [] }));
vi.mock("../../toast/toastStore", () => ({ pushToast: vi.fn() }));

const ANALISE: AnaliseInativacao = {
  usuarios: [
    {
      cpf: "12345678909",
      cpfMascarado: "***.456.789-**",
      nome: "Ana Souza",
      email: "ana@x.com",
      situacao: "EXECUTAVEL",
      alerta: null,
      estruturasViajante: ["APR001"],
      comoAprovador: [{ aprovacaoId: "APR011", posicoes: [1], segundoNivel: false, acao: "ORFA" }],
      candidatos: [],
    },
  ],
  resumo: { executaveis: 1, estruturasExcluidas: 1, estruturasCompactadas: 0, estruturasOrfas: 1, duplicados: [] },
  impressaoDigital: "digital-1",
  avisos: [],
};

const cadastro = new File(["c"], "cadastro.xlsx");
const estruturas = new File(["e"], "estruturas.xlsx");

async function ateAnalise() {
  const user = userEvent.setup();
  render(<InativacaoTab />);
  await user.upload(document.getElementById("inativacao_cadastro") as HTMLInputElement, cadastro);
  await user.upload(document.getElementById("inativacao_estruturas") as HTMLInputElement, estruturas);
  await user.type(document.getElementById("lista_text") as HTMLTextAreaElement, "12345678909");
  await user.click(document.getElementById("inativacao_btn") as HTMLElement);
  await screen.findByText("Conferir o impacto", { selector: "h3" });
  return user;
}

describe("InativacaoTab", () => {
  beforeEach(() => {
    vi.mocked(api.postAnalisar).mockResolvedValue(ANALISE);
    vi.stubGlobal("URL", Object.assign(URL, { createObjectURL: vi.fn(() => "blob:x"), revokeObjectURL: vi.fn() }));
  });
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("analisa, mostra o impacto com CPF mascarado e só executa com as duas confirmações", async () => {
    const user = await ateAnalise();
    expect(screen.getByText(/\*\*\*\.456\.789-\*\*/)).toBeInTheDocument();
    expect(screen.queryByText(/12345678909/, { selector: "li *" })).not.toBeInTheDocument();

    await user.click(document.getElementById("inativacao_next_btn") as HTMLElement);
    const executar = document.getElementById("inativacao_execute_btn") as HTMLButtonElement;
    expect(executar).toBeDisabled();
    await user.click(document.getElementById("inativacao_confirm_impacto") as HTMLElement);
    expect(executar).toBeDisabled();
    await user.click(document.getElementById("inativacao_confirm_orfas") as HTMLElement);
    expect(executar).toBeEnabled();
  });

  it("durante a execução trava tudo; concluída, só oferece Nova inativação", async () => {
    let terminar: (b: Blob) => void = () => {};
    vi.mocked(api.postExecutar).mockReturnValue(new Promise<Blob>((resolve) => (terminar = resolve)));
    const user = await ateAnalise();
    await user.click(document.getElementById("inativacao_next_btn") as HTMLElement);
    await user.click(document.getElementById("inativacao_confirm_impacto") as HTMLElement);
    await user.click(document.getElementById("inativacao_confirm_orfas") as HTMLElement);
    await user.click(document.getElementById("inativacao_execute_btn") as HTMLElement);

    expect(document.getElementById("inativacao_execute_btn")).toBeDisabled();
    expect(document.getElementById("inativacao_back_btn")).toBeDisabled();
    expect(document.getElementById("inativacao_confirm_impacto")).toBeDisabled();
    expect(document.getElementById("inativacao_confirm_orfas")).toBeDisabled();
    expect(document.getElementById("inativacao_new_btn")).toBeNull();

    terminar(new Blob(["zip"]));
    await waitFor(() => expect(document.getElementById("inativacao_status")).not.toBeNull());
    expect(document.getElementById("inativacao_new_btn")).toBeEnabled();
    expect(document.getElementById("inativacao_back_btn")).toBeNull();
    expect(document.getElementById("inativacao_execute_btn")).toBeNull();
  });

  it("avisa, sem bloquear, das linhas que não são CPF, e-mail nem nome completo", async () => {
    const user = userEvent.setup();
    render(<InativacaoTab />);
    expect(document.getElementById("lista_invalid_warning")).toBeNull();
    await user.type(document.getElementById("lista_text") as HTMLTextAreaElement, "1234567890{Enter}Maria{Enter}12345678909");
    const aviso = document.getElementById("lista_invalid_warning") as HTMLElement;
    expect(aviso).toHaveTextContent(/2 linhas ignoradas \(não são CPF de 11 dígitos, e-mail nem nome completo\): 1234567890, Maria/);
    expect(document.getElementById("lista_text")).toHaveAttribute("aria-describedby", expect.stringContaining("lista_invalid_warning"));
  });

  it("mostra a falha da análise na própria tela", async () => {
    vi.mocked(api.postAnalisar).mockRejectedValue(new api.InativacaoApiError("Base inválida", "ERRO_INTERNO"));
    const user = userEvent.setup();
    render(<InativacaoTab />);
    await user.upload(document.getElementById("inativacao_cadastro") as HTMLInputElement, cadastro);
    await user.upload(document.getElementById("inativacao_estruturas") as HTMLInputElement, estruturas);
    await user.type(document.getElementById("lista_text") as HTMLTextAreaElement, "12345678909");
    await user.click(document.getElementById("inativacao_btn") as HTMLElement);
    await waitFor(() => expect(document.getElementById("inativacao_debug")).not.toBeNull());
    expect(document.getElementById("inativacao_debug")).toHaveTextContent(/Base inválida/);
  });

  it("mostra o banner de confirmação do CPF do viajante e reanalisa ao confirmar", async () => {
    vi.mocked(api.postAnalisar)
      .mockRejectedValueOnce(
        new api.InativacaoApiError("CPF do viajante precisa de confirmação", "CPF_VIAJANTE_A_CONFIRMAR", {
          colunaCandidata: "Valor",
        }),
      )
      .mockResolvedValueOnce(ANALISE);
    const user = userEvent.setup();
    render(<InativacaoTab />);
    await user.upload(document.getElementById("inativacao_cadastro") as HTMLInputElement, cadastro);
    await user.upload(document.getElementById("inativacao_estruturas") as HTMLInputElement, estruturas);
    await user.type(document.getElementById("lista_text") as HTMLTextAreaElement, "12345678909");
    await user.click(document.getElementById("inativacao_btn") as HTMLElement);

    await waitFor(() => expect(document.getElementById("inativacao_cpf_viajante_confirm")).not.toBeNull());
    expect(document.getElementById("inativacao_cpf_viajante_confirm")).toHaveTextContent(/Valor/);
    const confirmBtn = document.getElementById("inativacao_confirm_cpf_viajante_btn") as HTMLElement;
    expect(confirmBtn).not.toBeNull();

    await user.click(confirmBtn);
    await waitFor(() => expect(document.getElementById("inativacao_cpf_viajante_confirm")).toBeNull());
    expect(api.postAnalisar).toHaveBeenLastCalledWith(cadastro, estruturas, ["12345678909"], [], true);
  });

  it("mostra os avisos de qualidade na etapa de impacto quando existem", async () => {
    vi.mocked(api.postAnalisar).mockResolvedValue({
      ...ANALISE,
      avisos: ["2 estruturas VIAJANTE sem CPF reconhecível na base"],
    });
    await ateAnalise();
    const lista = document.getElementById("inativacao_avisos") as HTMLElement;
    expect(lista).not.toBeNull();
    expect(lista).toHaveTextContent("2 estruturas VIAJANTE sem CPF reconhecível na base");
  });
});
