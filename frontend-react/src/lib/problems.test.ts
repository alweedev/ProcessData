import { describe, expect, it } from "vitest";
import type { LineDetail, QualityReport } from "./api";
import {
  coveredByGeneral,
  generalProblemText,
  generalProblems,
  lineDetails,
  problemTitle,
  summarizeProblems,
} from "./problems";

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

const detalhes: LineDetail[] = [
  {
    label: "Linha 3",
    nome: "ANA SOUZA",
    erros: ["Campo obrigatório em branco: Telefone"],
  },
  { label: "Linha 4", nome: "BRUNO LIMA", erros: ["CPF deve ter 11 dígitos"] },
  {
    label: "Linha 5",
    nome: "CARLA DIAS",
    erros: ["Campo obrigatório em branco: Telefone", "Email inválido"],
  },
];

describe("problemTitle — o problema em palavras de quem corrige a ficha", () => {
  it.each([
    ["Campo obrigatório em branco: Telefone", "Telefone em branco"],
    ["Campo obrigatório em branco: Centro de custo - Código", "Centro de custo - Código em branco"],
    ["CPF deve ter 11 dígitos", "CPF inválido (deve ter 11 dígitos)"],
    ["Email inválido", "E-mail inválido"],
    ["Nivel inválido, ajustado para vazio", "Nível inválido (ficou vazio)"],
    ["algo que o servidor passou a dizer", "algo que o servidor passou a dizer"], // desconhecido: mostra como veio
  ])("%s -> %s", (mensagem, titulo) => {
    expect(problemTitle(mensagem)).toBe(titulo);
  });
});

describe("problemas da ficha inteira", () => {
  it("coluna ausente e coluna vazia ganham frase clara", () => {
    expect(generalProblemText("Coluna obrigatória ausente na ficha: Telefone")).toBe(
      'A ficha não tem a coluna "Telefone".',
    );
    expect(generalProblemText("Coluna obrigatória vazia na ficha: E-mail")).toBe(
      'A coluna "E-mail" está vazia em todas as linhas.',
    );
  });

  it("separa vários e deixa passar o que não conhece (ex.: arquivo que não abriu)", () => {
    const report = relatorio({
      general_errors:
        "Coluna obrigatória ausente na ficha: Telefone; quebrada.xlsx: File contains no valid workbook part",
    });

    expect(generalProblems(report)).toEqual([
      'A ficha não tem a coluna "Telefone".',
      "quebrada.xlsx: File contains no valid workbook part",
    ]);
  });

  it("sem erro geral, lista vazia", () => {
    expect(generalProblems(relatorio())).toEqual([]);
  });
});

describe("summarizeProblems — o comum dito uma vez, cada passageiro só com a exceção", () => {
  const em = (label: string, nome: string, ...erros: string[]): LineDetail => ({
    label,
    nome,
    erros,
  });
  const branco = (...campos: string[]) => campos.map((c) => `Campo obrigatório em branco: ${c}`);

  it("o que falta em todas as linhas vai para 'comum'; cada linha guarda só o que tem a mais", () => {
    const resumo = summarizeProblems([
      em("Linha 3", "ANA SOUZA", ...branco("Telefone", "CPF")),
      em("Linha 4", "BRUNO LIMA", ...branco("Telefone", "CPF")),
      em("Linha 5", "CARLA DIAS", ...branco("Telefone", "CPF"), "Email inválido"),
    ]);

    expect(resumo.common).toEqual(["Telefone em branco", "CPF em branco"]);
    expect(resumo.rows).toEqual([
      { label: "Linha 3", nome: "ANA SOUZA", extras: [] },
      { label: "Linha 4", nome: "BRUNO LIMA", extras: [] },
      { label: "Linha 5", nome: "CARLA DIAS", extras: ["E-mail inválido"] },
    ]);
    expect(resumo.blankRows).toEqual([]);
  });

  it("problemas diferentes por linha: nada é comum e cada uma lista os seus", () => {
    const resumo = summarizeProblems(detalhes.slice(0, 2));

    expect(resumo.common).toEqual([]);
    expect(resumo.rows.map((r) => r.extras)).toEqual([["Telefone em branco"], ["CPF inválido (deve ter 11 dígitos)"]]);
  });

  it("uma linha só não tem 'comum': o problema fica na própria linha", () => {
    const resumo = summarizeProblems([detalhes[2]]);

    expect(resumo.common).toEqual([]);
    expect(resumo.rows[0].extras).toEqual(["Telefone em branco", "E-mail inválido"]);
  });

  it("linhas sem nenhum obrigatório ficam à parte e não entram na conta do comum", () => {
    const resumo = summarizeProblems([
      {
        ...em("Linha 2", "VAZIA UM", ...branco("CPF", "Telefone", "E-mail")),
        sem_preenchimento: true,
      },
      {
        ...em("Linha 3", "VAZIA DOIS", ...branco("CPF", "Telefone", "E-mail")),
        sem_preenchimento: true,
      },
      em("Linha 4", "ANA SOUZA", ...branco("Telefone")),
      em("Linha 5", "BRUNO LIMA", ...branco("Telefone", "E-mail")),
    ]);

    expect(resumo.blankRows).toEqual([
      { label: "Linha 2", nome: "VAZIA UM" },
      { label: "Linha 3", nome: "VAZIA DOIS" },
    ]);
    expect(resumo.common).toEqual(["Telefone em branco"]);
    expect(resumo.rows.map((r) => r.extras)).toEqual([[], ["E-mail em branco"]]);
  });

  it("o que um problema geral da ficha já explica não se repete por passageiro", () => {
    const resumo = summarizeProblems(
      [em("Linha 2", "ANA SOUZA", ...branco("Telefone")), em("Linha 3", "BRUNO LIMA", ...branco("Telefone", "E-mail"))],
      ["Telefone em branco"],
    );

    expect(resumo.common).toEqual([]);
    expect(resumo.rows).toEqual([{ label: "Linha 3", nome: "BRUNO LIMA", extras: ["E-mail em branco"] }]); // Ana só tinha o problema já explicado: some da lista
  });

  it("sem problemas, resumo vazio", () => {
    expect(summarizeProblems([])).toEqual({
      blankRows: [],
      common: [],
      rows: [],
    });
  });
});

describe("coveredByGeneral", () => {
  it("coluna ausente ou vazia cobre o 'em branco' daquela coluna; o resto não", () => {
    const report = relatorio({
      general_errors:
        "Coluna obrigatória ausente na ficha: Telefone; Coluna obrigatória vazia na ficha: E-mail; quebrada.xlsx: erro",
    });

    expect(coveredByGeneral(report)).toEqual(["Telefone em branco", "E-mail em branco"]);
    expect(coveredByGeneral(relatorio())).toEqual([]);
  });
});

describe("lineDetails", () => {
  it("usa line_details quando o servidor manda", () => {
    expect(lineDetails(relatorio({ line_details: detalhes }))).toBe(detalhes);
  });

  it("servidor antigo: cai no texto de line_errors, sem o nome do passageiro", () => {
    const report = relatorio({
      line_errors: {
        "Linha 3": "Campo obrigatório em branco: Telefone; Email inválido",
      },
    });

    expect(lineDetails(report)).toEqual([
      {
        label: "Linha 3",
        nome: "",
        erros: ["Campo obrigatório em branco: Telefone", "Email inválido"],
      },
    ]);
  });
});
