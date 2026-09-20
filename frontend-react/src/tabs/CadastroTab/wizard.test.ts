import { describe, expect, it } from "vitest";
import { estadosDaLinhaDoTempo, etapaMaxima, podeIrPara } from "./wizard";

type Pre = readonly [boolean, boolean, boolean];

describe("etapaMaxima — até onde dá para ir com o que já foi preenchido", () => {
  it.each<[Pre, number]>([
    [[false, false, false], 0],
    [[true, false, false], 1],
    [[true, true, false], 2],
    [[true, true, true], 3],
    // Removeu as fichas depois de escolher login e fluxo: só a etapa das fichas está liberada.
    [[false, true, true], 0],
    // Pulou o login mas escolheu o fluxo: para no login.
    [[true, false, true], 1],
  ])("%j -> etapa %i", (pre, esperado) => {
    expect(etapaMaxima(pre)).toBe(esperado);
  });
});

describe("podeIrPara", () => {
  const pre: Pre = [true, true, false]; // fichas e login ok, fluxo pendente -> alcance 2

  it("libera as etapas até o alcance e bloqueia as seguintes", () => {
    expect([0, 1, 2, 3].map((i) => podeIrPara(i, pre, false))).toEqual([true, true, true, false]);
  });

  it("depois de concluir o cadastro nenhuma etapa é clicável (só 'Novo cadastro' recomeça)", () => {
    expect([0, 1, 2, 3].map((i) => podeIrPara(i, [true, true, true], true))).toEqual([false, false, false, false]);
  });
});

describe("estadosDaLinhaDoTempo", () => {
  it("marca a etapa que está sendo vista como atual e as preenchidas como concluídas", () => {
    expect(estadosDaLinhaDoTempo(1, [true, false, false], false)).toEqual(["done", "current", "todo", "todo"]);
  });

  it("voltar a uma etapa já concluída a mostra como atual, sem apagar as outras concluídas", () => {
    expect(estadosDaLinhaDoTempo(0, [true, true, true], false)).toEqual(["current", "done", "done", "todo"]);
  });

  it("a etapa final só fica atual quando é a que está sendo vista", () => {
    expect(estadosDaLinhaDoTempo(3, [true, true, true], false)).toEqual(["done", "done", "done", "current"]);
  });

  it("cadastro concluído: todas as etapas concluídas, mesmo com os campos já zerados", () => {
    expect(estadosDaLinhaDoTempo(3, [false, false, false], true)).toEqual(["done", "done", "done", "done"]);
  });
});
