import { describe, expect, it } from "vitest";
import type { QualityReport } from "./api";
import { pendenciasDoRelatorio } from "./qualityReport";

const limpo: QualityReport = {
  total_rows: 10,
  valid_rows: 10,
  invalid_rows: 0,
  duplicated_rows: 0,
  general_errors: "",
  line_errors: {},
  required_blank: {},
};

describe("pendenciasDoRelatorio", () => {
  it("planilha limpa não tem pendência alguma", () => {
    expect(pendenciasDoRelatorio(limpo)).toEqual({ graves: [], avisos: [] });
  });

  it("linhas inválidas são graves, com singular e plural", () => {
    expect(pendenciasDoRelatorio({ ...limpo, invalid_rows: 1 }).graves).toEqual(["1 linha inválida"]);
    expect(pendenciasDoRelatorio({ ...limpo, invalid_rows: 3 }).graves).toEqual(["3 linhas inválidas"]);
  });

  it("erro geral é grave e vem primeiro", () => {
    expect(pendenciasDoRelatorio({ ...limpo, general_errors: "Coluna obrigatória ausente na ficha: Telefone", invalid_rows: 1 }).graves).toEqual([
      "erro geral na planilha",
      "1 linha inválida",
    ]);
  });

  it("linhas duplicadas são só aviso: o sistema já as tira do arquivo, não há o que confirmar", () => {
    expect(pendenciasDoRelatorio({ ...limpo, duplicated_rows: 1 })).toEqual({
      graves: [],
      avisos: ["1 linha duplicada removida"],
    });
    expect(pendenciasDoRelatorio({ ...limpo, duplicated_rows: 2 }).avisos).toEqual(["2 linhas duplicadas removidas"]);
  });

  it("campo obrigatório em branco não vira pendência à parte: já invalida a linha (o detalhe fica no relatório)", () => {
    expect(pendenciasDoRelatorio({ ...limpo, required_blank: { Telefone: 3 } })).toEqual({ graves: [], avisos: [] });
  });
});
