import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { UsuarioAnalise } from "../../lib/inativacaoApi";
import { ImpactoCard } from "./ImpactoCard";

function usuario(extra: Partial<UsuarioAnalise> = {}): UsuarioAnalise {
  return {
    cpf: "12345678909",
    cpfMascarado: "***.456.789-**",
    nome: "Ana Souza",
    email: "ana@x.com",
    situacao: "EXECUTAVEL",
    alerta: null,
    estruturasViajante: [],
    comoAprovador: [],
    candidatos: [],
    ...extra,
  };
}

function renderizar(u: UsuarioAnalise, escolhidos: string[] = [], onEscolher = vi.fn()) {
  render(
    <ul>
      <ImpactoCard usuario={u} escolhidos={escolhidos} onEscolher={onEscolher} />
    </ul>,
  );
  return onEscolher;
}

describe("ImpactoCard", () => {
  afterEach(cleanup);

  it("mostra o CPF mascarado e nunca o completo", () => {
    renderizar(usuario());
    expect(screen.getByText(/\*\*\*\.456\.789-\*\*/)).toBeInTheDocument();
    expect(screen.queryByText(/12345678909/)).not.toBeInTheDocument();
  });

  it("executável: mostra a estrutura do viajante a excluir e o impacto como aprovador", () => {
    renderizar(
      usuario({
        estruturasViajante: ["APR001"],
        comoAprovador: [
          { aprovacaoId: "APR010", posicoes: [2], segundoNivel: false, acao: "COMPACTACAO" },
          { aprovacaoId: "APR011", posicoes: [1], segundoNivel: true, acao: "ORFA" },
        ],
      }),
    );
    expect(screen.getByText(/Será excluída: APR001/)).toBeInTheDocument();
    expect(screen.getByText("Compactação")).toBeInTheDocument();
    expect(screen.getByText("Estrutura órfã")).toBeInTheDocument();
    expect(screen.getByText(/posição 2/)).toBeInTheDocument();
    expect(screen.getByText(/2º nível/)).toBeInTheDocument();
    expect(screen.getByText(/ficará sem nenhum aprovador/)).toBeInTheDocument();
  });

  it("executável sem estrutura nenhuma diz isso", () => {
    renderizar(usuario());
    expect(screen.getByText("Nenhuma estrutura direta encontrada.")).toBeInTheDocument();
    expect(screen.getByText("Não aparece como aprovador em outras estruturas.")).toBeInTheDocument();
  });

  it("sem CPF mostra o alerta e fica fora da cascata", () => {
    renderizar(usuario({ cpf: null, cpfMascarado: "", situacao: "SEM_CPF", alerta: "Sem CPF registrado." }));
    expect(screen.getByText("Sem CPF registrado.")).toBeInTheDocument();
    expect(screen.getByText("Sem CPF no cadastro")).toBeInTheDocument();
    expect(screen.queryByText("Compactação")).not.toBeInTheDocument();
  });

  it("homônimos: cada candidato é uma caixa de seleção; sem CPF não dá para escolher", async () => {
    const onEscolher = renderizar(
      usuario({
        cpf: null,
        situacao: "PENDENTE_SELECAO",
        candidatos: [
          { cpf: "11111111111", cpfMascarado: "***.111.111-**", nome: "João Silva", email: "j1@x.com" },
          { cpf: "", cpfMascarado: "", nome: "João Silva", email: "j2@x.com" },
        ],
      }),
    );
    const [comCpf, semCpf] = screen.getAllByRole("checkbox");
    expect(semCpf).toBeDisabled();
    await userEvent.click(comCpf);
    expect(onEscolher).toHaveBeenCalledWith("11111111111", true);
  });

  it("desabilitado: nenhum candidato pode ser marcado", () => {
    render(
      <ul>
        <ImpactoCard
          usuario={usuario({
            cpf: null,
            situacao: "PENDENTE_SELECAO",
            candidatos: [{ cpf: "11111111111", cpfMascarado: "***.111.111-**", nome: "João Silva", email: "j1@x.com" }],
          })}
          escolhidos={[]}
          onEscolher={vi.fn()}
          disabled
        />
      </ul>,
    );
    expect(screen.getByRole("checkbox")).toBeDisabled();
  });
});
