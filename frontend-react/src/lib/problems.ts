import type { LineDetail, QualityReport } from "./api";

export interface ProblemRow {
  label: string;
  nome: string;
}

/** Uma linha com problema, já sem o que é comum a todas: só o que ela tem a mais. */
export interface ProblemRowSummary extends ProblemRow {
  extras: string[];
}

/** O que a tela mostra: o problema repetido em todas as linhas dito uma vez, e cada passageiro só com a sua exceção. */
export interface ProblemSummary {
  /** linhas sem nenhum campo obrigatório: parecem não preenchidas (não entram na conta do que é comum) */
  blankRows: ProblemRow[];
  /** problemas que TODAS as demais linhas têm (só existe com 2 linhas ou mais) */
  common: string[];
  rows: ProblemRowSummary[];
}

const BLANK_FIELD = /^Campo obrigatório em branco: (.+)$/;

/** O problema de uma linha em palavras de quem vai corrigir a ficha ("Telefone em branco"), não do validador. */
export function problemTitle(message: string): string {
  const blank = BLANK_FIELD.exec(message);
  if (blank) return `${blank[1]} em branco`;
  if (message.startsWith("CPF deve ter")) return "CPF inválido (deve ter 11 dígitos)";
  if (message === "Email inválido") return "E-mail inválido";
  if (message.startsWith("Nivel inválido")) return "Nível inválido (ficou vazio)";
  return message;
}

/** Problema da ficha inteira (coluna que falta, coluna toda vazia, arquivo que não abriu). */
export function generalProblemText(part: string): string {
  const missing = /^Coluna obrigatória ausente na ficha: (.+)$/.exec(part);
  if (missing) return `A ficha não tem a coluna "${missing[1]}".`;
  const empty = /^Coluna obrigatória vazia na ficha: (.+)$/.exec(part);
  if (empty) return `A coluna "${empty[1]}" está vazia em todas as linhas.`;
  return part;
}

export function generalProblems(report: QualityReport): string[] {
  return report.general_errors ? report.general_errors.split("; ").filter(Boolean).map(generalProblemText) : [];
}

/** As linhas com problema. Servidor antigo (sem `line_details`) cai no texto de `line_errors`, sem o nome do passageiro. */
export function lineDetails(report: QualityReport): LineDetail[] {
  if (report.line_details) return report.line_details;
  return Object.entries(report.line_errors).map(([label, message]) => ({
    label,
    nome: "",
    erros: message.split("; "),
  }));
}

/** Os problemas por linha que um problema geral da ficha já explica: coluna ausente/vazia => "X em branco" em toda
 *  linha. Repeti-los por passageiro seria dizer a mesma coisa duas vezes. */
export function coveredByGeneral(report: QualityReport): string[] {
  const covered: string[] = [];
  for (const part of report.general_errors ? report.general_errors.split("; ") : []) {
    const column = /^Coluna obrigatória (?:ausente|vazia) na ficha: (.+)$/.exec(part);
    if (column) covered.push(`${column[1]} em branco`);
  }
  return covered;
}

/** "Comum + exceções": em vez de repetir o mesmo problema e o nome de cada passageiro por problema, diz uma vez o que
 *  falta em todas as linhas e, por passageiro, só o que difere. Linhas sem nenhum obrigatório (parecem não preenchidas)
 *  ficam à parte para não fingir que cada uma tem 7 problemas diferentes. */
export function summarizeProblems(details: LineDetail[], covered: string[] = []): ProblemSummary {
  const blankRows = details.filter((d) => d.sem_preenchimento).map(({ label, nome }) => ({ label, nome }));
  const partial = details
    .filter((d) => !d.sem_preenchimento)
    .map((d) => ({
      label: d.label,
      nome: d.nome,
      titles: [...new Set(d.erros.map(problemTitle))].filter((title) => !covered.includes(title)),
    }))
    .filter((row) => row.titles.length > 0);

  const common =
    partial.length < 2 ? [] : partial[0].titles.filter((title) => partial.every((row) => row.titles.includes(title)));
  const rows = partial.map(({ label, nome, titles }) => ({
    label,
    nome,
    extras: titles.filter((title) => !common.includes(title)),
  }));
  return { blankRows, common, rows };
}
