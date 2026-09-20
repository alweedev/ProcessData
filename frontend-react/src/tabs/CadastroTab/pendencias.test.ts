import { describe, expect, it } from "vitest";
import { descreverPendencias } from "./pendencias";

describe("descreverPendencias", () => {
  it.each([
    [{ temArquivos: true, login: "CPF", fluxo: "SELF" }, null],
    [
      { temArquivos: false, login: null, fluxo: null },
      "Para continuar: enviar ao menos uma ficha e escolher o tipo de login e o fluxo.",
    ],
    [{ temArquivos: false, login: "CPF", fluxo: "SELF" }, "Para continuar: enviar ao menos uma ficha."],
    [{ temArquivos: true, login: null, fluxo: null }, "Para continuar: escolher o tipo de login e o fluxo."],
    [{ temArquivos: true, login: null, fluxo: "FRONT" }, "Para continuar: escolher o tipo de login."],
    [{ temArquivos: true, login: "EMAIL", fluxo: null }, "Para continuar: escolher o fluxo."],
    [
      { temArquivos: false, login: "CPF", fluxo: null },
      "Para continuar: enviar ao menos uma ficha e escolher o fluxo.",
    ],
  ])("%j -> %j", (entrada, esperado) => {
    expect(descreverPendencias(entrada.temArquivos, entrada.login, entrada.fluxo)).toBe(esperado);
  });
});
