import { cleanup, render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it } from "vitest";
import type { QualityReport } from "../lib/api";
import { ValidationReport } from "./ValidationReport";

const relatorio = (extra: Partial<QualityReport> = {}): QualityReport => ({
  total_rows: 5,
  valid_rows: 1,
  invalid_rows: 4,
  duplicated_rows: 0,
  general_errors: "",
  line_errors: {},
  required_blank: {},
  ...extra,
});

const comProblemas = relatorio({
  line_details: [
    {
      label: "Linha 3",
      nome: "ANA SOUZA",
      erros: ["Campo obrigatório em branco: Telefone"],
    },
    {
      label: "Linha 4",
      nome: "BRUNO LIMA",
      erros: ["CPF deve ter 11 dígitos"],
    },
    {
      label: "Linha 5",
      nome: "CARLA DIAS",
      erros: ["Campo obrigatório em branco: Telefone", "Email inválido"],
    },
  ],
});

describe("ValidationReport — o detalhe das pendências", () => {
  afterEach(cleanup);

  it("planilha sem problema não mostra nada (o veredito é o 'Tudo certo')", () => {
    const { container } = render(<ValidationReport report={relatorio({ invalid_rows: 0 })} />);

    expect(container).toBeEmptyDOMElement();
  });

  it("sem relatório não mostra nada", () => {
    const { container } = render(<ValidationReport report={null} />);

    expect(container).toBeEmptyDOMElement();
  });

  it("cada passageiro aparece uma vez; sem problema em comum, ao lado do nome vai o que falta", () => {
    render(<ValidationReport report={comProblemas} />);

    expect(screen.queryByRole("heading")).not.toBeInTheDocument(); // sem título nem contagem: o cartão de cima já diz
    expect(screen.queryByText(/^Em todos|^Nos dois/)).not.toBeInTheDocument();
    expect(screen.getAllByText(/ANA SOUZA|BRUNO LIMA|CARLA DIAS/)).toHaveLength(3);
    const carla = screen.getByText("CARLA DIAS").closest("li") as HTMLElement;
    expect(carla).toHaveTextContent("Linha 5 · CARLA DIAS — Telefone em branco; E-mail inválido");
  });

  it("o que falta em todos é dito uma vez, e cada passageiro só traz a exceção", () => {
    const branco = ["Campo obrigatório em branco: Telefone", "Campo obrigatório em branco: E-mail"];
    render(
      <ValidationReport
        report={relatorio({
          line_details: [
            { label: "Linha 3", nome: "ANA SOUZA", erros: branco },
            { label: "Linha 4", nome: "BRUNO LIMA", erros: branco },
            {
              label: "Linha 5",
              nome: "CARLA DIAS",
              erros: [...branco, "CPF deve ter 11 dígitos"],
            },
          ],
        })}
      />,
    );

    expect(screen.getByText("Em todos os 3:")).toBeInTheDocument();
    expect(screen.getAllByText("Telefone em branco")).toHaveLength(1);
    expect(screen.getByText("Linha 3").closest("li")).toHaveTextContent("Linha 3 · ANA SOUZA");
    expect(screen.getByText("Linha 3").closest("li")).not.toHaveTextContent("também");
    expect(screen.getByText("CARLA DIAS").closest("li")).toHaveTextContent(
      "— também: CPF inválido (deve ter 11 dígitos)",
    );
  });

  it("linhas sem nenhum obrigatório: só diz isso, sem listar os problemas de cada uma", () => {
    const todos = ["CPF", "Telefone", "E-mail"].map((c) => `Campo obrigatório em branco: ${c}`);
    render(
      <ValidationReport
        report={relatorio({
          line_details: [
            {
              label: "Linha 2",
              nome: "VAZIA UM",
              erros: todos,
              sem_preenchimento: true,
            },
            {
              label: "Linha 3",
              nome: "VAZIA DOIS",
              erros: todos,
              sem_preenchimento: true,
            },
          ],
        })}
      />,
    );

    const bloco = document.getElementById("cadastro_problems_blank") as HTMLElement;
    expect(bloco).toHaveTextContent("2 linhas sem nenhum campo obrigatório — parece que não foram preenchidas");
    expect(bloco).toHaveTextContent("Linha 2 · VAZIA UM");
    expect(screen.queryByText(/em branco/)).not.toBeInTheDocument();
    expect(document.getElementById("cadastro_problems_rows")).toBeNull();
  });

  it("linha sem nome mostra só o número da linha", () => {
    const report = relatorio({
      line_details: [{ label: "Linha 7", nome: "", erros: ["Email inválido"] }],
    });
    render(<ValidationReport report={report} />);

    expect(screen.getByText("Linha 7")).toBeInTheDocument();
    expect(screen.queryByText(/·/)).not.toBeInTheDocument();
  });

  it("coluna ausente: a frase da ficha basta, sem repetir 'em branco' de cada passageiro", () => {
    render(
      <ValidationReport
        report={relatorio({
          general_errors: "Coluna obrigatória ausente na ficha: Telefone",
          line_details: [
            {
              label: "Linha 2",
              nome: "ANA SOUZA",
              erros: ["Campo obrigatório em branco: Telefone"],
            },
            {
              label: "Linha 3",
              nome: "BRUNO LIMA",
              erros: ["Campo obrigatório em branco: Telefone"],
            },
          ],
        })}
      />,
    );

    expect(screen.getByText('A ficha não tem a coluna "Telefone".')).toBeInTheDocument();
    expect(screen.queryByText(/em branco/)).not.toBeInTheDocument();
    expect(screen.queryByText(/ANA SOUZA/)).not.toBeInTheDocument();
  });

  it("problema da ficha inteira aparece em destaque, em palavras claras", () => {
    render(
      <ValidationReport
        report={relatorio({
          invalid_rows: 0,
          general_errors: "Coluna obrigatória ausente na ficha: Telefone",
        })}
      />,
    );

    expect(screen.getByText('A ficha não tem a coluna "Telefone".')).toBeInTheDocument();
  });

  it("lista grande mostra 5 passageiros e abre o resto sob demanda", async () => {
    const muitas = Array.from({ length: 8 }, (_, i) => ({
      label: `Linha ${i + 2}`,
      nome: `PESSOA ${i}`,
      erros: ["Campo obrigatório em branco: Telefone"],
    }));
    render(<ValidationReport report={relatorio({ line_details: muitas })} />);

    const bloco = document.getElementById("cadastro_problems_rows") as HTMLElement;
    expect(within(bloco).getAllByRole("listitem")).toHaveLength(1 + 5); // o problema comum + 5 passageiros

    await userEvent.click(within(bloco).getByRole("button", { name: "Mostrar todos (3 a mais)" }));
    expect(within(bloco).getAllByRole("listitem")).toHaveLength(1 + 8);

    await userEvent.click(within(bloco).getByRole("button", { name: "Mostrar menos" }));
    expect(within(bloco).getAllByRole("listitem")).toHaveLength(1 + 5);
  });

  it("erro de validação (servidor recusou, rede) aparece como mensagem", () => {
    render(<ValidationReport report={null} error="Não foi possível validar: A planilha está vazia." />);

    expect(screen.getByText("Não foi possível validar: A planilha está vazia.")).toBeInTheDocument();
  });
});
