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

  it("voltar a uma etapa a mostra como atual e troca o verde das seguintes por 'preenchida' (valor guardado)", () => {
    expect(estadosDaLinhaDoTempo(0, [true, true, true], false)).toEqual(["current", "ready", "ready", "todo"]);
    expect(estadosDaLinhaDoTempo(1, [true, true, true], false)).toEqual(["done", "current", "ready", "todo"]);
    expect(estadosDaLinhaDoTempo(2, [true, true, true], false)).toEqual(["done", "done", "current", "todo"]);
  });

  it("'preenchida' só vale com tudo antes preenchido: fluxo escolhido com o login vazio continua pendente", () => {
    expect(estadosDaLinhaDoTempo(0, [true, true, false], false)).toEqual(["current", "ready", "todo", "todo"]);
    expect(estadosDaLinhaDoTempo(0, [true, false, true], false)).toEqual(["current", "todo", "todo", "todo"]);
  });

  it("sem fichas nenhuma etapa seguinte fica concluída, mesmo com login e fluxo guardados", () => {
    expect(estadosDaLinhaDoTempo(0, [false, true, true], false)).toEqual(["current", "todo", "todo", "todo"]);
    // Defensivo: mesmo que a tela esteja numa etapa adiante, o que não tem base não conta como concluído.
    expect(estadosDaLinhaDoTempo(2, [false, true, true], false)).toEqual(["todo", "todo", "current", "todo"]);
  });

  it("a etapa final só fica atual quando é a que está sendo vista", () => {
    expect(estadosDaLinhaDoTempo(3, [true, true, true], false)).toEqual(["done", "done", "done", "current"]);
  });

  it("cadastro concluído: todas as etapas concluídas, mesmo com os campos já zerados", () => {
    expect(estadosDaLinhaDoTempo(3, [false, false, false], true)).toEqual(["done", "done", "done", "done"]);
  });
});
