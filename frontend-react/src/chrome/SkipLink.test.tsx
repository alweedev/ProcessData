import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it } from "vitest";
import { SkipLink } from "./SkipLink";

describe("SkipLink", () => {
  afterEach(cleanup);

  it("leva o foco do teclado direto ao conteúdo, sem mexer no hash da URL (o roteador é o hash)", async () => {
    window.location.hash = "#/cadastro";
    render(
      <>
        <SkipLink />
        <main id="conteudo" tabIndex={-1}>
          conteúdo
        </main>
      </>,
    );

    await userEvent.click(screen.getByRole("button", { name: "Pular para o conteúdo" }));

    expect(document.getElementById("conteudo")).toHaveFocus();
    expect(window.location.hash).toBe("#/cadastro");
  });
});
