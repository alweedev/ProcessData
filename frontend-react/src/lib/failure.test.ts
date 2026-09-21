import { describe, expect, it } from "vitest";
import { explainFailure } from "./failure";

const NO_COLUMNS =
  "Nenhuma coluna da ficha foi reconhecida (colunas lidas: Operacao, AprovacaoId, Valor). O cabeçalho precisa estar na 1ª linha.";

describe("explainFailure — a falha da geração em palavras de quem corrige a ficha", () => {
  it("arquivo que não é ficha: motivo por arquivo, colunas lidas à parte e o que fazer", () => {
    const view = explainFailure({
      message: `Nenhum registro processado: mapa.xls: ${NO_COLUMNS}`,
      fileErrors: { "mapa.xls": NO_COLUMNS },
    });

    expect(view.title).toBe("Não foi possível gerar o cadastro");
    expect(view.causes).toEqual([
      {
        file: "mapa.xls",
        text: "Não parece uma ficha de cadastro: nenhuma coluna conhecida foi encontrada.",
        detail: "Colunas lidas: Operacao, AprovacaoId, Valor",
      },
    ]);
    expect(view.tips).toContain("O cabeçalho precisa estar na 1ª linha da planilha.");
    expect(view.technical).toContain("Nenhum registro processado: mapa.xls");
  });

  it("a dica de várias fichas com o mesmo problema não se repete", () => {
    const view = explainFailure({
      message: "x",
      fileErrors: { "a.xls": NO_COLUMNS, "b.xls": NO_COLUMNS },
    });

    expect(view.causes.map((c) => c.file)).toEqual(["a.xls", "b.xls"]);
    expect(view.tips.filter((t) => t.startsWith("O cabeçalho"))).toHaveLength(1);
  });

  it("outras abas: avisa que só a 1ª é lida", () => {
    const view = explainFailure({
      message: "x",
      fileErrors: { "a.xls": `${NO_COLUMNS} Só a 1ª aba ('Dados') é lida.` },
    });

    expect(view.tips).toContain('Só a 1ª aba ("Dados") é lida: deixe os dados nela.');
  });

  it("sem separação por arquivo, tira o prefixo técnico e mantém a frase do servidor", () => {
    const view = explainFailure({ message: "Nenhum registro processado: nenhuma linha com dados foi encontrada." });

    expect(view.causes).toEqual([
      { file: undefined, text: "nenhuma linha com dados foi encontrada.", detail: undefined },
    ]);
    expect(view.technical).toBeUndefined(); // não reescreveu nada: não há o que mostrar como "técnico"
  });

  it("nomes acima de 20 caracteres: manda para a conferência de nomes", () => {
    const view = explainFailure({
      message: "2 nomes com mais de 20 caracteres em Nome ou Sobrenome (Linha 3, Linha 4). Ajuste na conferência.",
    });

    expect(view.tips).toEqual(["Ajuste esses nomes na conferência de nomes e gere de novo."]);
  });

  it.each([
    ["Erro de rede.", "Não conseguimos falar com o servidor."],
    ["Erro 500", "O servidor teve um problema (erro 500)."],
  ])("falha de rede/servidor (%s)", (message, texto) => {
    const view = explainFailure({ message });

    expect(view.causes[0].text).toBe(texto);
    expect(view.tips).toHaveLength(1);
    expect(view.technical).toBe(message);
  });

  it("o título muda conforme a etapa (validar ou gerar)", () => {
    expect(explainFailure({ message: "x" }).title).toBe("Não foi possível gerar o cadastro");
    expect(explainFailure({ message: "x" }, "Não foi possível validar a planilha").title).toBe(
      "Não foi possível validar a planilha",
    );
  });

  it("mensagem desconhecida aparece como veio", () => {
    expect(explainFailure({ message: "algo novo do servidor" }).causes[0].text).toBe("algo novo do servidor");
  });
});
