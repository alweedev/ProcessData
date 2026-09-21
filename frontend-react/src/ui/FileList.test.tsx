import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import { FileList } from "./FileList";

function fakeFile(name: string, sizeBytes: number): File {
  const file = new File(["x"], name);
  Object.defineProperty(file, "size", { value: sizeBytes });
  return file;
}

describe("FileList", () => {
  afterEach(cleanup);

  it("não renderiza nada sem arquivos", () => {
    const { container } = render(<FileList files={[]} maxFiles={5} onRemove={() => {}} />);

    expect(container).toBeEmptyDOMElement();
  });

  it("resume a quantidade em relação ao limite e o tamanho total (1 KB + 2 KB = 3.0 KB)", () => {
    render(<FileList files={[fakeFile("a.xlsx", 1024), fakeFile("b.xlsx", 2048)]} maxFiles={5} onRemove={() => {}} />);

    expect(screen.getByText(/2 de 5/)).toBeInTheDocument();
    expect(screen.getByText(/arquivos · 3\.0 KB no total/)).toBeInTheDocument();
  });

  it("usa o singular com um único arquivo", () => {
    render(<FileList files={[fakeFile("a.xlsx", 1024)]} maxFiles={5} onRemove={() => {}} />);

    expect(screen.getByText(/1 de 5/)).toBeInTheDocument();
    expect(screen.getByText(/arquivo · 1\.0 KB no total/)).toBeInTheDocument();
  });

  it("remover chama onRemove com o índice do arquivo certo", async () => {
    const onRemove = vi.fn();
    render(<FileList files={[fakeFile("a.xlsx", 1), fakeFile("b.xlsx", 1)]} maxFiles={5} onRemove={onRemove} />);

    await userEvent.click(screen.getByRole("button", { name: "Remover b.xlsx" }));

    expect(onRemove).toHaveBeenCalledExactlyOnceWith(1);
  });

  describe("efeito de entrada", () => {
    const linhas = () => screen.getAllByRole("listitem");
    const props = { maxFiles: 5, onRemove: () => {} };

    it("ficha adicionada depois entra animada, escalonada; a que já estava ali não repete", () => {
      const a = fakeFile("a.xlsx", 1);
      const b = fakeFile("b.xlsx", 1);
      const c = fakeFile("c.xlsx", 1);
      const { rerender } = render(<FileList files={[]} {...props} />);

      rerender(<FileList files={[a]} {...props} />);
      expect(linhas()[0]).toHaveAttribute("data-entering");

      rerender(<FileList files={[a, b, c]} {...props} />);
      const [la, lb, lc] = linhas();
      expect(la).toHaveAttribute("data-entering"); // já entrou animada; continua marcada, sem reiniciar
      expect(lb.style.getPropertyValue("--d")).toBe("0ms");
      expect(lc.style.getPropertyValue("--d")).toBe("80ms");
    });

    it("remover uma ficha não reanima as que ficaram", () => {
      const a = fakeFile("a.xlsx", 1);
      const b = fakeFile("b.xlsx", 1);
      const { rerender } = render(<FileList files={[a, b]} {...props} />);
      const linhaB = linhas()[1];

      rerender(<FileList files={[b]} {...props} />);

      expect(linhas()[0]).toBe(linhaB); // mesma linha no DOM: chave estável, nada remontou
    });

    it("ao montar já com fichas (voltar à etapa) nenhuma anima", () => {
      render(<FileList files={[fakeFile("a.xlsx", 1), fakeFile("b.xlsx", 1)]} {...props} />);

      for (const linha of linhas()) expect(linha).not.toHaveAttribute("data-entering");
    });
  });

  it("'Limpar todos' só existe quando há onClearAll e o aciona", async () => {
    const onClearAll = vi.fn();
    const { rerender } = render(<FileList files={[fakeFile("a.xlsx", 1)]} maxFiles={5} onRemove={() => {}} />);
    expect(screen.queryByRole("button", { name: "Remover todos os arquivos" })).not.toBeInTheDocument();

    rerender(<FileList files={[fakeFile("a.xlsx", 1)]} maxFiles={5} onRemove={() => {}} onClearAll={onClearAll} />);
    await userEvent.click(screen.getByRole("button", { name: "Remover todos os arquivos" }));

    expect(onClearAll).toHaveBeenCalledOnce();
  });
});
