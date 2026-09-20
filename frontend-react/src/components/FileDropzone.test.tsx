import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import { FileDropzone } from "./FileDropzone";

type Variant = "classic" | "rich";

function setup(props: { variant?: Variant; compact?: boolean; onFiles?: (f: FileList) => void } = {}) {
  const onFiles = props.onFiles ?? vi.fn();
  render(
    <FileDropzone
      id="arquivos"
      containerId="area"
      accept=".xlsx"
      ariaLabel="Upload"
      description="Arraste aqui"
      hint="Excel até 10 MB"
      onFiles={onFiles}
      {...props}
    />,
  );
  return { area: document.getElementById("area") as HTMLElement, onFiles };
}

/** O jsdom não tem DragEvent: um MouseEvent carrega o `relatedTarget` de verdade,
 *  que é o que o componente lê para distinguir "saiu da área" de "entrou num filho". */
function dragLeave(area: HTMLElement, relatedTarget: Element) {
  fireEvent(area, new MouseEvent("dragleave", { bubbles: true, relatedTarget }));
}

describe("FileDropzone", () => {
  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
  });

  describe("variante rich", () => {
    it("clicar em qualquer ponto da área abre o seletor de arquivos", async () => {
      const abrir = vi.spyOn(HTMLInputElement.prototype, "click").mockImplementation(() => {});
      setup({ variant: "rich" });

      await userEvent.click(screen.getByText("Arraste aqui"));

      expect(abrir).toHaveBeenCalledOnce();
    });

    it("clicar no botão Selecionar abre o seletor uma única vez (o clique não é tratado duas vezes)", async () => {
      const abrir = vi.spyOn(HTMLInputElement.prototype, "click").mockImplementation(() => {});
      setup({ variant: "rich" });

      await userEvent.click(screen.getByRole("button", { name: "Selecionar" }));

      expect(abrir).toHaveBeenCalledOnce();
    });

    it("mostra 'Solte para enviar' ao arrastar sobre a área e volta ao texto normal ao sair", () => {
      const { area } = setup({ variant: "rich" });

      fireEvent.dragOver(area);
      expect(screen.getByText("Solte para enviar")).toBeInTheDocument();

      dragLeave(area, document.body);
      expect(screen.getByText("Arraste aqui")).toBeInTheDocument();
    });

    it("passar o arrasto por cima de um elemento interno não faz o destaque piscar", () => {
      const { area } = setup({ variant: "rich" });
      fireEvent.dragOver(area);

      dragLeave(area, screen.getByRole("button", { name: "Selecionar" }));

      expect(screen.getByText("Solte para enviar")).toBeInTheDocument();
    });

    it("soltar arquivos os entrega a onFiles e encerra o destaque", () => {
      const { area, onFiles } = setup({ variant: "rich" });
      const arquivo = new File(["x"], "fichas.xlsx");

      fireEvent.dragOver(area);
      fireEvent.drop(area, { dataTransfer: { files: [arquivo] } });

      expect(onFiles).toHaveBeenCalledOnce();
      expect(vi.mocked(onFiles).mock.calls[0][0][0]).toBe(arquivo);
      expect(screen.getByText("Arraste aqui")).toBeInTheDocument();
    });

    it("no modo compacto esconde a dica de formatos", () => {
      setup({ variant: "rich", compact: true });

      expect(screen.queryByText("Excel até 10 MB")).not.toBeInTheDocument();
    });

    it("fora do modo compacto mostra a dica de formatos", () => {
      setup({ variant: "rich" });

      expect(screen.getByText("Excel até 10 MB")).toBeInTheDocument();
    });
  });

  describe("variante classic (Inativação e Estruturas, ainda sem o redesign)", () => {
    it("só o botão abre o seletor: clicar no texto da área não faz nada", async () => {
      const abrir = vi.spyOn(HTMLInputElement.prototype, "click").mockImplementation(() => {});
      setup({ variant: "classic" });

      await userEvent.click(screen.getByText("Arraste aqui"));
      expect(abrir).not.toHaveBeenCalled();

      await userEvent.click(screen.getByRole("button", { name: "Selecionar" }));
      expect(abrir).toHaveBeenCalledOnce();
    });

    it("não mostra a dica de formatos (recurso só da variante rich)", () => {
      setup({ variant: "classic" });

      expect(screen.queryByText("Excel até 10 MB")).not.toBeInTheDocument();
    });
  });

  it("Enter na área abre o seletor (acionamento por teclado)", async () => {
    const abrir = vi.spyOn(HTMLInputElement.prototype, "click").mockImplementation(() => {});
    const { area } = setup({ variant: "rich" });
    area.focus();

    await userEvent.keyboard("{Enter}");

    expect(abrir).toHaveBeenCalledOnce();
  });
});
