import { describe, expect, it } from "vitest";
import { normalizeCpf } from "./useEstruturas";

describe("normalizeCpf", () => {
  it("remove pontuação, mantendo só dígitos", () => {
    expect(normalizeCpf("123.456.789-00")).toBe("12345678900");
  });

  it("string já normalizada permanece igual", () => {
    expect(normalizeCpf("12345678900")).toBe("12345678900");
  });

  it("string vazia ou undefined vira string vazia", () => {
    expect(normalizeCpf("")).toBe("");
    expect(normalizeCpf(undefined as unknown as string)).toBe("");
  });
});
