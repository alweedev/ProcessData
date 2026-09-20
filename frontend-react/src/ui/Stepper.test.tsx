import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
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

  describe("navegação pelos pontos (onSelect)", () => {
    const navegaveis = [
      { label: "Enviar", state: "done" as const, selectable: true },
      { label: "Configurar", state: "current" as const, selectable: true },
      { label: "Gerar", state: "todo" as const },
    ];

    it("etapas ao alcance (fora a atual) viram botões e chamam onSelect com o índice", async () => {
      const onSelect = vi.fn();
      render(<Stepper label="Etapas" steps={navegaveis} onSelect={onSelect} />);

      await userEvent.click(screen.getByRole("button", { name: /Enviar/ }));

      expect(onSelect).toHaveBeenCalledExactlyOnceWith(0);
    });

    it("a etapa atual e as fora de alcance não são botões", () => {
      render(<Stepper label="Etapas" steps={navegaveis} onSelect={() => {}} />);

      expect(screen.queryByRole("button", { name: /Configurar/ })).not.toBeInTheDocument();
      expect(screen.queryByRole("button", { name: /Gerar/ })).not.toBeInTheDocument();
    });

    it("sem onSelect a linha do tempo é só informativa: nenhum botão", () => {
      render(<Stepper label="Etapas" steps={navegaveis} />);

      expect(screen.queryAllByRole("button")).toHaveLength(0);
    });
  });

  it("mostra a linha de detalhe da etapa quando ela existe", () => {
    render(<Stepper label="Etapas" steps={steps} />);

    expect(screen.getByText("2 arquivos")).toBeInTheDocument();
  });
});
