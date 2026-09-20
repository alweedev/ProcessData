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

  it("outras teclas não mudam a seleção", async () => {
    render(<Harness />);
    screen.getByRole("radio", { name: "CPF" }).focus();

    await userEvent.keyboard("a");

    expect(screen.getByRole("radio", { name: "CPF" })).toHaveAttribute("aria-checked", "true");
  });
});
