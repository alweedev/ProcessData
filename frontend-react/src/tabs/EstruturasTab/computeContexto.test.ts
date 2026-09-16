import { describe, expect, it } from "vitest";
import { computeContexto } from "./computeContexto";
import type { PreviewItem } from "./useEstruturas";

function item(overrides: Partial<PreviewItem>): PreviewItem {
  return { aprovacaoId: "A1", aprovacaoPor: "VIAJANTE", ...overrides };
}

describe("computeContexto", () => {
  it("VIAJANTE com nome retorna o nome do viajante", () => {
    expect(computeContexto(item({ aprovacaoPor: "VIAJANTE", viajanteNomeCompleto: "Ana Souza" }))).toBe("Ana Souza");
  });

  it("CCEMPRESA com código e descrição retorna 'código - descrição'", () => {
    expect(computeContexto(item({ aprovacaoPor: "CCEMPRESA", ccCodigo: "CC001", ccDescricao: "TI" }))).toBe(
      "CC001 - TI",
    );
  });

  it("CCEMPRESA só com código retorna o código", () => {
    expect(computeContexto(item({ aprovacaoPor: "CCEMPRESA", ccCodigo: "CC001" }))).toBe("CC001");
  });

  it("CCEMPRESA sem código nem descrição retorna vazio", () => {
    expect(computeContexto(item({ aprovacaoPor: "CCEMPRESA" }))).toBe("");
  });

  it("tipo desconhecido cai no fallback: prioriza código+descrição sobre nome", () => {
    expect(
      computeContexto(
        item({ aprovacaoPor: "OUTRO", ccCodigo: "CC1", ccDescricao: "Desc", viajanteNomeCompleto: "Ana" }),
      ),
    ).toBe("CC1 - Desc");
  });

  it("tipo desconhecido sem código/descrição usa o nome do viajante", () => {
    expect(computeContexto(item({ aprovacaoPor: "OUTRO", viajanteNomeCompleto: "Ana" }))).toBe("Ana");
  });

  it("VIAJANTE sem nome mas com CC cai no fallback genérico", () => {
    expect(computeContexto(item({ aprovacaoPor: "VIAJANTE", ccCodigo: "CC1", ccDescricao: "Desc" }))).toBe(
      "CC1 - Desc",
    );
  });

  it("nenhum dado disponível retorna string vazia", () => {
    expect(computeContexto(item({ aprovacaoPor: "OUTRO" }))).toBe("");
  });
});
