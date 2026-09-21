import { describe, expect, it } from "vitest";
import type { NameReviewItem } from "./api";
import {
  acceptAll,
  applyEdit,
  buildOverrides,
  decisionFor,
  fits,
  initialDecision,
  isResolved,
  pendingItems,
  sanitizeName,
} from "./nameReview";

const duvidoso: NameReviewItem = {
  key: "1:2",
  label: "Linha 2",
  nome_completo: "XANTO ZEKO QUIM BRAG",
  nome: "XANTO",
  sobrenome: "ZEKO QUIM BRAG",
  confianca: "baixa",
  motivos: ["a divisão entre nome e sobrenome é ambígua"],
  estouro: { nome: false, sobrenome: false },
  sugestao: null,
};

const estourado: NameReviewItem = {
  key: "1:3",
  label: "Linha 3",
  nome_completo: "ANNA LUIZA FERNANDES DE ALBUQUERQUE CAVALCANTI",
  nome: "ANNA LUIZA",
  sobrenome: "FERNANDES DE ALBUQUERQUE CAVALCANTI",
  confianca: "alta",
  motivos: [],
  estouro: { nome: false, sobrenome: true },
  sugestao: { nome: "ANNA LUIZA", sobrenome: "F DE A CAVALCANTI" },
};

const semSolucao: NameReviewItem = {
  ...estourado,
  key: "1:4",
  sugestao: { nome: "ANNA LUIZA", sobrenome: "SOBRENOMEMUITOMUITOLONGOSOZINHO" },
};

describe("sanitizeName — a mesma limpeza do backend", () => {
  it("maiúsculo, sem acento, sem apóstrofo, espaços compactados", () => {
    expect(sanitizeName("  José   D'Ávila  Nandú ")).toBe("JOSE DAVILA NANDU");
  });

  it("mantém hífen, barra e parênteses", () => {
    expect(sanitizeName("Carlos-Eduardo (Dudu) Lima/Filho")).toBe("CARLOS-EDUARDO (DUDU) LIMA/FILHO");
  });
});

describe("fits", () => {
  it("cabe com até 20 caracteres em cada campo", () => {
    expect(fits("A".repeat(20), "B".repeat(20))).toBe(true);
    expect(fits("A".repeat(21), "B")).toBe(false);
    expect(fits("A", "B".repeat(21))).toBe(false);
  });

  it("conta o texto depois da limpeza, não o digitado", () => {
    expect(fits("áéíóú ".repeat(3), "silva")).toBe(true); // 15 letras + espaços, acentos não contam a mais
  });

  it("os dois precisam estar preenchidos", () => {
    expect(fits("", "SILVA")).toBe(false);
    expect(fits("ANA", "   ")).toBe(false);
  });
});

describe("decisões da conferência", () => {
  it("parte da divisão sugerida, ou da abreviação quando o nome passou de 20", () => {
    expect(initialDecision(duvidoso)).toEqual({ nome: "XANTO", sobrenome: "ZEKO QUIM BRAG", confirmed: false, edited: false });
    expect(initialDecision(estourado).sobrenome).toBe("F DE A CAVALCANTI");
  });

  it("nada decidido = pendente; a decisão padrão vem do item", () => {
    expect(pendingItems([duvidoso, estourado], {})).toHaveLength(2);
    expect(decisionFor(duvidoso, {})).toEqual(initialDecision(duvidoso));
  });

  it("corrigir à mão confirma sozinho quando cabe e marca como editado", () => {
    const decisao = applyEdit(initialDecision(duvidoso), { nome: "XANTO ZEKO", sobrenome: "QUIM BRAG" });

    expect(decisao).toMatchObject({ nome: "XANTO ZEKO", sobrenome: "QUIM BRAG", edited: true, confirmed: true });
    expect(isResolved(decisao)).toBe(true);
  });

  it("corrigir para algo que não cabe deixa pendente", () => {
    const decisao = applyEdit(initialDecision(duvidoso), { sobrenome: "S".repeat(25) });

    expect(decisao.confirmed).toBe(false);
    expect(isResolved(decisao)).toBe(false);
  });

  it("'Aceitar todas' confirma o que cabe e deixa pendente o que não cabe nem abreviado", () => {
    const decisoes = acceptAll([duvidoso, estourado, semSolucao], {});

    expect(pendingItems([duvidoso, estourado, semSolucao], decisoes).map((i) => i.key)).toEqual(["1:4"]);
    expect(decisoes["1:2"].edited).toBe(false); // aceitar não é editar
  });

  it("uma confirmação que deixou de caber volta a pendente", () => {
    const decisao = { ...initialDecision(duvidoso), confirmed: true, sobrenome: "S".repeat(25) };

    expect(isResolved(decisao)).toBe(false);
  });
});

describe("buildOverrides — o que vai junto na geração", () => {
  it("só as decisões resolvidas, com o nome completo de referência e o texto já limpo", () => {
    const decisoes = {
      "1:2": applyEdit(initialDecision(duvidoso), { nome: "Xanto  Zêko", sobrenome: "Quim Brag" }),
      "1:3": { ...initialDecision(estourado), confirmed: true },
    };

    expect(buildOverrides([duvidoso, estourado, semSolucao], decisoes)).toEqual({
      "1:2": { nome_completo: "XANTO ZEKO QUIM BRAG", nome: "XANTO ZEKO", sobrenome: "QUIM BRAG", editado: true },
      "1:3": {
        nome_completo: "ANNA LUIZA FERNANDES DE ALBUQUERQUE CAVALCANTI",
        nome: "ANNA LUIZA",
        sobrenome: "F DE A CAVALCANTI",
        editado: false,
      },
    });
  });

  it("nada resolvido, nada a enviar", () => {
    expect(buildOverrides([duvidoso], {})).toEqual({});
  });
});
