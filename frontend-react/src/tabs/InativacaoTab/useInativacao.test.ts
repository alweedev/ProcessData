import { describe, expect, it } from "vitest";
import { classifyList, formatCpf, isValidFullName } from "./useInativacao";

describe("isValidFullName", () => {
  it("aceita nome com 2+ partes e 3+ caracteres", () => {
    expect(isValidFullName("João Silva")).toBe(true);
  });

  it("rejeita uma única palavra", () => {
    expect(isValidFullName("Maria")).toBe(false);
  });

  it("caso limite: 2 partes de 1 caractere cada já soma os 3 caracteres mínimos", () => {
    // Não dá pra ter 2+ palavras com menos de 3 caracteres no total (a
    // checagem de comprimento é redundante com a de partes, mas documentamos
    // o comportamento real em vez de assumir um caso que não existe).
    expect(isValidFullName("A B")).toBe(true);
  });

  it("ignora acentos na contagem de caracteres", () => {
    expect(isValidFullName("Ana Léa")).toBe(true);
  });
});

describe("classifyList", () => {
  it("separa CPF, nome e e-mail em listas distintas", () => {
    const result = classifyList("123.456.789-00\nJoão Silva\nfulano@empresa.com");
    expect(result.validCpfs).toEqual(["12345678900"]);
    expect(result.validNames).toEqual(["João Silva"]);
    expect(result.validEmails).toEqual(["fulano@empresa.com"]);
    expect(result.totalValid).toBe(3);
  });

  it("detecta CPF duplicado (com e sem formatação) e conta só 1 válido", () => {
    const result = classifyList("12345678900\n123.456.789-00");
    expect(result.validCpfs).toEqual(["12345678900"]);
    expect(result.duplicates).toEqual(["12345678900"]);
    expect(result.totalValid).toBe(1);
  });

  it("descarta nome inválido (uma palavra só) sem quebrar a classificação", () => {
    const result = classifyList("Maria\nJoão Silva");
    expect(result.validNames).toEqual(["João Silva"]);
    expect(result.totalValid).toBe(1);
  });

  it("ignora linhas em branco", () => {
    const result = classifyList("\n\n12345678900\n\n");
    expect(result.totalValid).toBe(1);
  });

  it("texto vazio não gera itens válidos nem quebra", () => {
    expect(classifyList("")).toEqual({
      validCpfs: [],
      validNames: [],
      validEmails: [],
      duplicates: [],
      totalValid: 0,
    });
  });
});

describe("formatCpf", () => {
  it("formata 11 dígitos como XXX.XXX.XXX-XX", () => {
    expect(formatCpf("12345678900")).toBe("123.456.789-00");
  });

  it("devolve a entrada sem alteração se não tiver 11 dígitos", () => {
    expect(formatCpf("123")).toBe("123");
  });

  it("string vazia vira string vazia", () => {
    expect(formatCpf("")).toBe("");
  });
});
