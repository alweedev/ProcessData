import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { ProcessingProgress } from "./ProcessingProgress";

describe("ProcessingProgress", () => {
  afterEach(cleanup);

  it("durante o envio mostra a porcentagem real", () => {
    render(<ProcessingProgress id="p" progress={40} />);
    const bar = screen.getByRole("progressbar");
    expect(bar).toHaveAttribute("aria-valuenow", "40");
    expect(screen.getByText("Enviando fichas…")).toBeInTheDocument();
    expect(screen.getByText("40%")).toBeInTheDocument();
  });

  it("depois do envio (100%) vira indeterminada: o servidor ainda processa, não há porcentagem", () => {
    render(<ProcessingProgress id="p" progress={100} />);
    const bar = screen.getByRole("progressbar");
    expect(bar).not.toHaveAttribute("aria-valuenow");
    expect(screen.getByText("Processando no servidor…")).toBeInTheDocument();
    expect(screen.queryByText("100%")).not.toBeInTheDocument();
  });
});
