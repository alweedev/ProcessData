import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it } from "vitest";
import type { QualityReport } from "../lib/api";
import { ValidationSummary } from "./ValidationSummary";

const relatorio = (extra: Partial<QualityReport> = {}): QualityReport => ({
  total_rows: 45,
  valid_rows: 45,
  invalid_rows: 0,
  duplicated_rows: 0,
  general_errors: "",
  line_errors: {},
  required_blank: {},
  ...extra,
});

describe("ValidationSummary", () => {
  afterEach(cleanup);

  it("enquanto valida mostra o andamento", () => {
    render(<ValidationSummary status="loading" report={null} />);

    expect(screen.getByRole("status")).toHaveTextContent("Validando a planilha…");
  });

  describe("tudo certo: uma linha só", () => {
    it("diz quantos cadastros estão prontos, sem repetir a mesma coisa em outros textos", () => {
      render(<ValidationSummary status="done" report={relatorio()} />);

      const resumo = screen.getByRole("status");
      expect(resumo).toHaveTextContent("Tudo certo: 45 cadastros prontos para gerar.");
      expect(resumo).toHaveAttribute("data-tone", "success");
      expect(resumo).not.toHaveTextContent("Nenhum problema");
      expect(resumo).not.toHaveTextContent("informativa");
    });

    it("no singular", () => {
      render(<ValidationSummary status="done" report={relatorio({ total_rows: 1, valid_rows: 1 })} />);

      expect(screen.getByRole("status")).toHaveTextContent("Tudo certo: 1 cadastro pronto para gerar.");
    });

    it("os números ficam atrás de 'Ver detalhes' e abrem só a pedido", async () => {
      render(<ValidationSummary status="done" report={relatorio()} />);
      expect(screen.queryByText("Linhas")).not.toBeInTheDocument();

      const alternar = screen.getByRole("button", { name: "Ver detalhes" });
      expect(alternar).toHaveAttribute("aria-expanded", "false");
      await userEvent.click(alternar);

      expect(screen.getByRole("button", { name: "Ocultar detalhes" })).toHaveAttribute("aria-expanded", "true");
      for (const rotulo of ["Linhas", "Válidas", "Inválidas", "Duplicadas"])
        expect(screen.getByText(rotulo)).toBeInTheDocument();

      await userEvent.click(screen.getByRole("button", { name: "Ocultar detalhes" }));
      expect(screen.queryByText("Linhas")).not.toBeInTheDocument();
    });
  });

  it("só aviso (duplicadas removidas): segue 'Tudo certo', diz o que foi feito e não promete confirmação", () => {
    render(
      <ValidationSummary
        status="done"
        report={relatorio({
          total_rows: 43,
          valid_rows: 43,
          duplicated_rows: 2,
        })}
      />,
    );

    const resumo = screen.getByRole("status");
    expect(resumo).toHaveTextContent(
      "Tudo certo, com um aviso: 2 linhas duplicadas removidas. 43 cadastros prontos para gerar.",
    );
    expect(resumo).not.toHaveTextContent("confirmação");
    expect(resumo).toHaveAttribute("data-tone", "info");
    expect(screen.getByRole("button", { name: "Ver detalhes" })).toBeInTheDocument();
  });

  describe("com pendências", () => {
    const comPendencias = relatorio({
      total_rows: 5,
      valid_rows: 1,
      invalid_rows: 4,
      duplicated_rows: 1,
    });

    it("diz quantos cadastros têm pendência, o que já está pronto e o que fazer, sem repetir a contagem", () => {
      render(<ValidationSummary status="done" report={comPendencias} />);

      const resumo = screen.getByRole("status");
      expect(resumo).toHaveTextContent("4 de 5 cadastros têm pendência.");
      expect(resumo).toHaveTextContent("1 pronto · 1 linha repetida removida (a planilha tinha 6)");
      expect(resumo).not.toHaveTextContent("com pendência ·"); // o título já diz 4 de 5
      expect(resumo).toHaveTextContent("Corrija a ficha ou gere assim mesmo (pediremos sua confirmação).");
      expect(resumo).toHaveAttribute("data-tone", "warning");
    });

    it("o detalhe (children) entra no mesmo cartão; sem pendência ele não aparece", () => {
      const { rerender } = render(
        <ValidationSummary status="done" report={comPendencias}>
          <p>detalhe das pendências</p>
        </ValidationSummary>,
      );
      expect(screen.getByRole("status")).toContainElement(screen.getByText("detalhe das pendências"));

      rerender(
        <ValidationSummary status="done" report={relatorio()}>
          <p>detalhe das pendências</p>
        </ValidationSummary>,
      );
      expect(screen.queryByText("detalhe das pendências")).not.toBeInTheDocument();
    });

    it("não tem 'Ver detalhes': o detalhe é o painel de problemas", () => {
      render(<ValidationSummary status="done" report={comPendencias} />);

      expect(screen.queryByRole("button", { name: /detalhes/ })).not.toBeInTheDocument();
    });

    it.each([
      [relatorio({ total_rows: 5, valid_rows: 4, invalid_rows: 1 }), "1 de 5 cadastros tem pendência."],
      [relatorio({ total_rows: 1, valid_rows: 0, invalid_rows: 1 }), "O cadastro tem pendência."],
    ])("concordância certa no singular", (report, frase) => {
      render(<ValidationSummary status="done" report={report} />);

      expect(screen.getByRole("status")).toHaveTextContent(frase);
    });

    it("só problema geral da ficha (nenhuma linha inválida)", () => {
      render(
        <ValidationSummary
          status="done"
          report={relatorio({
            general_errors: "Coluna obrigatória ausente na ficha: Telefone",
          })}
        />,
      );

      expect(screen.getByRole("status")).toHaveTextContent("A planilha tem um problema geral.");
    });
  });

  it.each(["idle", "error"] as const)("em '%s' não mostra nada (o erro aparece no relatório)", (status) => {
    const { container } = render(<ValidationSummary status={status} report={null} />);

    expect(container).toBeEmptyDOMElement();
  });
});
