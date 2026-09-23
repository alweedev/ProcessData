import { describe, expect, it } from "vitest";
import type { AnaliseInativacao, Situacao, UsuarioAnalise } from "../../lib/inativacaoApi";
import { podeContinuar, podeExecutar } from "./wizard";

function usuario(situacao: Situacao): UsuarioAnalise {
  return {
    cpf: "12345678909",
    cpfMascarado: "***.456.789-**",
    nome: "Ana Souza",
    email: "ana@x.com",
    situacao,
    alerta: null,
    estruturasViajante: [],
    comoAprovador: [],
    candidatos: [],
  };
}

function analise(usuarios: UsuarioAnalise[], orfas = 0): AnaliseInativacao {
  return {
    usuarios,
    resumo: {
      executaveis: usuarios.filter((u) => u.situacao === "EXECUTAVEL").length,
      estruturasExcluidas: 0,
      estruturasCompactadas: 0,
      estruturasOrfas: orfas,
      duplicados: [],
    },
    impressaoDigital: "d",
    avisos: [],
  };
}

describe("podeContinuar", () => {
  it("exige análise com ao menos um executável", () => {
    expect(podeContinuar(null)).toBe(false);
    expect(podeContinuar(analise([usuario("SEM_CPF")]))).toBe(false);
    expect(podeContinuar(analise([usuario("EXECUTAVEL")]))).toBe(true);
  });

  it("bloqueia enquanto houver homônimo sem escolha", () => {
    expect(podeContinuar(analise([usuario("EXECUTAVEL"), usuario("PENDENTE_SELECAO")]))).toBe(false);
  });
});

describe("podeExecutar", () => {
  it("exige a confirmação do impacto", () => {
    const a = analise([usuario("EXECUTAVEL")]);
    expect(podeExecutar(a, { impacto: false, orfas: false })).toBe(false);
    expect(podeExecutar(a, { impacto: true, orfas: false })).toBe(true);
  });

  it("com estrutura órfã exige também a ciência dela", () => {
    const a = analise([usuario("EXECUTAVEL")], 2);
    expect(podeExecutar(a, { impacto: true, orfas: false })).toBe(false);
    expect(podeExecutar(a, { impacto: true, orfas: true })).toBe(true);
  });

  it("sem análise nunca executa", () => {
    expect(podeExecutar(null, { impacto: true, orfas: true })).toBe(false);
  });
});
