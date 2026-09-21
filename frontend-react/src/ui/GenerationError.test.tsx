import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { explainFailure } from "../lib/failure";
import { GenerationError } from "./GenerationError";

const NO_COLUMNS =
  "Nenhuma coluna da ficha foi reconhecida (colunas lidas: Operacao, Valor). O cabeçalho precisa estar na 1ª linha.";
const view = explainFailure({
  message: `Nenhum registro processado: mapa.xls: ${NO_COLUMNS}`,
  fileErrors: { "mapa.xls": NO_COLUMNS },
});

describe("GenerationError", () => {
  afterEach(cleanup);

  it("mostra o título, o motivo por arquivo, o que fazer e os detalhes técnicos recolhidos", () => {
    render(<GenerationError view={view} />);

    const caixa = document.getElementById("cadastro_debug") as HTMLElement;
    expect(caixa).toHaveAttribute("aria-live", "assertive");
    expect(caixa).toHaveTextContent("Não foi possível gerar o cadastro");
    expect(caixa).toHaveTextContent("mapa.xls: Não parece uma ficha de cadastro");
    expect(caixa).toHaveTextContent("Colunas lidas: Operacao, Valor");
    expect(screen.getByText("O que fazer")).toBeInTheDocument();
    expect(screen.getByText("Detalhes técnicos").closest("details")).not.toHaveAttribute("open");
  });

  it("o id do cartão pode mudar (validação e geração não repetem o mesmo id)", () => {
    render(<GenerationError view={view} id="cadastro_validation_error" />);

    expect(document.getElementById("cadastro_validation_error")).toBeInTheDocument();
    expect(document.getElementById("cadastro_debug")).toBeNull();
  });
});
