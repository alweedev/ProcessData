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

  describe("etapa preenchida adiante da atual ('ready')", () => {
    const comPreenchida = [
      { label: "Enviar", state: "current" as const },
      { label: "Configurar", state: "ready" as const },
      { label: "Gerar", state: "todo" as const },
    ];

    it("anuncia '(preenchida)', mantém o número e usa o contorno tracejado, sem virar ✓ de concluída", () => {
      render(<Stepper label="Etapas" steps={comPreenchida} />);

      expect(screen.getByText("(preenchida)")).toBeInTheDocument();
      expect(screen.getByTestId("stepper-dot-1")).toHaveTextContent("2");
      expect(screen.getByTestId("stepper-dot-1").className).toContain("border-dashed");
    });

    it("a linha só se preenche até o que foi alcançado: 'ready' não avança o progresso", () => {
      const { container } = render(<Stepper label="Etapas" steps={comPreenchida} />);

      const [linhaAteConfigurar] = container.querySelectorAll("li > span[aria-hidden] > span");
      expect(linhaAteConfigurar.className).toContain("scale-x-0");
    });
  });

  it("o anel se centra no ponto da etapa atual nos dois eixos, mesmo com o ponto fora do topo da lista", () => {
    // jsdom não faz layout: simula a lista em (100, 50) e o ponto atual 2px abaixo do topo dela
    // (é o que acontece quando o texto ao lado é mais alto que o ponto).
    const spy = vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockImplementation(function (this: HTMLElement) {
      const isDot = this.dataset.testid === "stepper-dot-1";
      const box = isDot ? { left: 300, top: 52, width: 32, height: 32 } : { left: 100, top: 50, width: 800, height: 36 };
      return { ...box, right: box.left + box.width, bottom: box.top + box.height, x: box.left, y: box.top, toJSON: () => ({}) };
    });
    try {
      render(<Stepper label="Etapas" steps={steps} />);

      // centro do ponto na lista: x = 300-100+16 = 216, y = 52-50+16 = 18; o anel (40px) sai de (196, -2).
      expect(screen.getByTestId("stepper-marker").style.transform).toBe("translate(196px, -2px)");
    } finally {
      spy.mockRestore();
    }
  });

  it("mostra a linha de detalhe da etapa quando ela existe", () => {
    render(<Stepper label="Etapas" steps={steps} />);

    expect(screen.getByText("2 arquivos")).toBeInTheDocument();
  });
});
