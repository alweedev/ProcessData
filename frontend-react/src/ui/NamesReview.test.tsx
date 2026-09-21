import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { NameReviewItem } from "../lib/api";
import { initialDecision, type Decisions } from "../lib/nameReview";
import { NamesReview } from "./NamesReview";

const duvidoso: NameReviewItem = {
  key: "1:2",
  label: "Linha 2",
  nome_completo: "XANTO ZEKO QUIM BRAG",
  nome: "XANTO",
  sobrenome: "ZEKO QUIM BRAG",
  confianca: "baixa",
  motivos: ["a divisão entre nome e sobrenome é ambígua"],
  estouro: { nome: false, sobrenome: false },
  sugestao: null,
};

const estourado: NameReviewItem = {
  key: "1:3",
  label: "Linha 3",
  nome_completo: "ANNA LUIZA FERNANDES DE ALBUQUERQUE CAVALCANTI",
  nome: "ANNA LUIZA",
  sobrenome: "FERNANDES DE ALBUQUERQUE CAVALCANTI",
  confianca: "alta",
  motivos: [],
  estouro: { nome: false, sobrenome: true },
  sugestao: { nome: "ANNA LUIZA", sobrenome: "F DE A CAVALCANTI" },
};

function setup(items: NameReviewItem[], decisions: Decisions = {}) {
  const onDecision = vi.fn();
  const onAcceptAll = vi.fn();
  render(<NamesReview items={items} decisions={decisions} onDecision={onDecision} onAcceptAll={onAcceptAll} />);
  return { onDecision, onAcceptAll };
}

describe("NamesReview", () => {
  afterEach(cleanup);

  it("lista cada nome com a linha, o nome do documento e o motivo da dúvida", () => {
    setup([duvidoso]);

    const linha = screen.getByRole("listitem", { name: /Linha 2/ });
    expect(within(linha).getByText("XANTO ZEKO QUIM BRAG")).toBeInTheDocument();
    expect(within(linha).getByText("a divisão entre nome e sobrenome é ambígua")).toBeInTheDocument();
    expect(screen.getByLabelText("Nome")).toHaveValue("XANTO");
    expect(screen.getByLabelText("Sobrenome")).toHaveValue("ZEKO QUIM BRAG");
  });

  it("diz quantos ainda precisam de confirmação e, resolvidos todos, libera", () => {
    const { rerender } = render(
      <NamesReview items={[duvidoso, estourado]} decisions={{}} onDecision={() => {}} onAcceptAll={() => {}} />,
    );
    expect(screen.getByRole("status")).toHaveTextContent("2 nomes ainda precisam da sua confirmação antes de gerar.");

    const confirmados: Decisions = {
      "1:2": { ...initialDecision(duvidoso), confirmed: true },
      "1:3": { ...initialDecision(estourado), confirmed: true },
    };
    rerender(<NamesReview items={[duvidoso, estourado]} decisions={confirmados} onDecision={() => {}} onAcceptAll={() => {}} />);

    expect(screen.getByRole("status")).toHaveTextContent("Todos os nomes foram conferidos. Pode gerar.");
    expect(screen.getAllByText("Conferido")).toHaveLength(2);
  });

  it("no singular", () => {
    setup([duvidoso]);

    expect(screen.getByRole("status")).toHaveTextContent("1 nome ainda precisa da sua confirmação");
  });

  it("nome que passou de 20 já vem com a abreviação sugerida e o contador certo", () => {
    setup([estourado]);

    expect(screen.getByLabelText("Sobrenome")).toHaveValue("F DE A CAVALCANTI");
    expect(screen.getByText("17/20")).toBeInTheDocument();
  });

  it("corrigir o texto avisa a nova decisão (editada e, se cabe, já confirmada)", () => {
    const { onDecision } = setup([duvidoso]);

    fireEvent.change(screen.getByLabelText("Nome"), { target: { value: "XANTO ZEKO" } });

    expect(onDecision).toHaveBeenCalledWith("1:2", {
      nome: "XANTO ZEKO",
      sobrenome: "ZEKO QUIM BRAG",
      confirmed: true,
      edited: true,
    });
  });

  it("texto acima de 20 mostra o erro no campo e não deixa confirmar", () => {
    setup([duvidoso], { "1:2": { ...initialDecision(duvidoso), sobrenome: "S".repeat(25) } });

    expect(screen.getByText("25/20: passa do limite")).toBeInTheDocument();
    expect(screen.getByLabelText("Sobrenome")).toHaveAttribute("aria-invalid", "true");
    expect(screen.getByRole("button", { name: "Confirmar" })).toBeDisabled();
    expect(screen.getByText(/Passa de 20 caracteres/)).toBeInTheDocument();
  });

  it("Confirmar aceita a sugestão como está (sem marcar como editada)", async () => {
    const { onDecision } = setup([duvidoso]);

    await userEvent.click(screen.getByRole("button", { name: "Confirmar" }));

    expect(onDecision).toHaveBeenCalledWith("1:2", { ...initialDecision(duvidoso), confirmed: true });
  });

  it("'Aceitar todas as sugestões' chama a ação e some da vez quando não há mais o que aceitar", async () => {
    const { onAcceptAll } = setup([duvidoso]);
    await userEvent.click(screen.getByRole("button", { name: "Aceitar todas as sugestões" }));
    expect(onAcceptAll).toHaveBeenCalledOnce();

    cleanup();
    setup([duvidoso], { "1:2": { ...initialDecision(duvidoso), confirmed: true } });
    expect(screen.getByRole("button", { name: "Aceitar todas as sugestões" })).toBeDisabled();
  });
});
