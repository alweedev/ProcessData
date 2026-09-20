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

  it("'Limpar todos' só existe quando há onClearAll e o aciona", async () => {
    const onClearAll = vi.fn();
    const { rerender } = render(<FileList files={[fakeFile("a.xlsx", 1)]} maxFiles={5} onRemove={() => {}} />);
    expect(screen.queryByRole("button", { name: "Remover todos os arquivos" })).not.toBeInTheDocument();

    rerender(<FileList files={[fakeFile("a.xlsx", 1)]} maxFiles={5} onRemove={() => {}} onClearAll={onClearAll} />);
    await userEvent.click(screen.getByRole("button", { name: "Remover todos os arquivos" }));

    expect(onClearAll).toHaveBeenCalledOnce();
  });
});
