import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { Stepper } from "./Stepper";

describe("Stepper", () => {
  afterEach(cleanup);

  const steps = [
    { label: "Enviar", detail: "2 arquivos", state: "done" as const },
    { label: "Configurar", state: "current" as const },
    { label: "Gerar", state: "todo" as const },
  ];

  it("marca só a etapa atual com aria-current e anuncia o estado das demais para leitores de tela", () => {
    render(<Stepper label="Etapas" steps={steps} />);
    const [enviar, configurar, gerar] = screen.getAllByRole("listitem");

    expect(configurar).toHaveAttribute("aria-current", "step");
    expect(enviar).not.toHaveAttribute("aria-current");
    expect(gerar).not.toHaveAttribute("aria-current");
    expect(screen.getByText("(concluída)")).toBeInTheDocument();
    expect(screen.getByText("(etapa atual)")).toBeInTheDocument();
  });

  it("etapa concluída troca o número pelo ✓; atual e pendente mantêm o número da posição", () => {
    render(<Stepper label="Etapas" steps={steps} />);

    expect(screen.queryByText("1")).not.toBeInTheDocument();
    expect(screen.getByText("2")).toBeInTheDocument();
    expect(screen.getByText("3")).toBeInTheDocument();
  });

  it("mostra a linha de detalhe da etapa quando ela existe", () => {
    render(<Stepper label="Etapas" steps={steps} />);

    expect(screen.getByText("2 arquivos")).toBeInTheDocument();
  });
});
