import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { useState } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { SegmentedControl } from "./SegmentedControl";

const OPTIONS = [
  { value: "CPF", label: "CPF" },
  { value: "EMAIL", label: "E-mail" },
  { value: "OUTRO", label: "Outro" },
];

function Harness({ initial = "CPF" }: { initial?: string }) {
  const [value, setValue] = useState(initial);
  return <SegmentedControl id="login" label="Tipo de login" value={value} options={OPTIONS} onChange={setValue} />;
}

describe("SegmentedControl", () => {
  afterEach(cleanup);

  it("expõe um radiogroup nomeado e só a opção marcada entra na ordem de Tab", () => {
    render(<Harness />);

    expect(screen.getByRole("radiogroup", { name: "Tipo de login" })).toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("aria-checked", "true");
    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("tabindex", "0");
    expect(screen.getByRole("radio", { name: "E-mail" })).toHaveAttribute("aria-checked", "false");
    expect(screen.getByRole("radio", { name: "E-mail" })).toHaveAttribute("tabindex", "-1");
  });

  it("clicar numa opção a seleciona", async () => {
    render(<Harness />);

    await userEvent.click(screen.getByRole("radio", { name: "E-mail" }));

    expect(screen.getByRole("radio", { name: "E-mail" })).toHaveAttribute("aria-checked", "true");
    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("aria-checked", "false");
  });

  it("→ avança e ← volta, levando o foco junto com a seleção", async () => {
    render(<Harness />);
    screen.getByRole("radio", { name: "CPF" }).focus();

    await userEvent.keyboard("{ArrowRight}");
    expect(screen.getByRole("radio", { name: "E-mail" })).toHaveAttribute("aria-checked", "true");
    expect(screen.getByRole("radio", { name: "E-mail" })).toHaveFocus();

    await userEvent.keyboard("{ArrowLeft}");
    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("aria-checked", "true");
    expect(screen.getByRole("radio", { name: "CPF" })).toHaveFocus();
  });

  it("a navegação por setas é circular nas duas pontas", async () => {
    render(<Harness initial="OUTRO" />);
    screen.getByRole("radio", { name: "Outro" }).focus();

    await userEvent.keyboard("{ArrowRight}");
    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("aria-checked", "true");

    await userEvent.keyboard("{ArrowLeft}");
    expect(screen.getByRole("radio", { name: "Outro" })).toHaveAttribute("aria-checked", "true");
  });

  describe("sem opção escolhida (value nulo)", () => {
    function Vazio() {
      const [value, setValue] = useState<string | null>(null);
      return <SegmentedControl id="login" label="Tipo de login" value={value} options={OPTIONS} onChange={setValue} />;
    }

    it("nenhuma opção vem marcada e a primeira continua alcançável por Tab", () => {
      render(<Vazio />);
      const radios = screen.getAllByRole("radio");

      expect(radios.map((r) => r.getAttribute("aria-checked"))).toEqual(["false", "false", "false"]);
      expect(radios.map((r) => r.getAttribute("tabindex"))).toEqual(["0", "-1", "-1"]);
    });

    it("com `required`, avisa que a escolha é obrigatória até alguém escolher", async () => {
      function Obrigatorio() {
        const [value, setValue] = useState<string | null>(null);
        return (
          <SegmentedControl id="login" label="Tipo de login" value={value} options={OPTIONS} onChange={setValue} required />
        );
      }
      render(<Obrigatorio />);

      expect(screen.getByRole("radiogroup", { name: /Tipo de login/ })).toHaveAttribute("aria-required", "true");
      expect(screen.getByText("obrigatório")).toBeInTheDocument();

      await userEvent.click(screen.getByRole("radio", { name: "CPF" }));

      expect(screen.queryByText("obrigatório")).not.toBeInTheDocument();
    });

    it("→ a partir da primeira opção seleciona a segunda", async () => {
      render(<Vazio />);
      screen.getByRole("radio", { name: "CPF" }).focus();

      await userEvent.keyboard("{ArrowRight}");

      expect(screen.getByRole("radio", { name: "E-mail" })).toHaveAttribute("aria-checked", "true");
    });
  });

  it("outras teclas não mudam a seleção", async () => {
    render(<Harness />);
    screen.getByRole("radio", { name: "CPF" }).focus();

    await userEvent.keyboard("a");

    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("aria-checked", "true");
  });
});
