import type { NameReviewItem } from "./api";

/** Limite de Nome e de Sobrenome na plataforma (o mesmo do backend: `MAX_FIELD_LEN`). */
export const MAX_NAME_FIELD = 20;

export interface NameDecision {
  nome: string;
  sobrenome: string;
  /** o usuário conferiu esta divisão (aceitou ou corrigiu) */
  confirmed: boolean;
  /** o usuário mudou o texto à mão (vale mais para o vocabulário do que só aceitar) */
  edited: boolean;
}

/** Mesma limpeza que o backend faz (maiúsculo, sem acento, só letras/números/espaço/`-/()`), para o contador
 *  de caracteres bater com o que vai para a plataforma. O backend continua sendo a palavra final. */
export function sanitizeName(text: string): string {
  return text
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9 \-/()]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/** Cabe na plataforma: os dois preenchidos e cada um com até 20 caracteres. */
export function fits(nome: string, sobrenome: string): boolean {
  const first = sanitizeName(nome);
  const last = sanitizeName(sobrenome);
  return first.length > 0 && last.length > 0 && first.length <= MAX_NAME_FIELD && last.length <= MAX_NAME_FIELD;
}

/** O ponto de partida do item: a divisão sugerida, ou a abreviação proposta quando passou de 20. */
export function initialDecision(item: NameReviewItem): NameDecision {
  const base = item.sugestao ?? { nome: item.nome, sobrenome: item.sobrenome };
  return { nome: base.nome, sobrenome: base.sobrenome, confirmed: false, edited: false };
}

export type Decisions = Record<string, NameDecision>;

export function decisionFor(item: NameReviewItem, decisions: Decisions): NameDecision {
  return decisions[item.key] ?? initialDecision(item);
}

/** Corrigir à mão já é uma decisão: confirma sozinho quando o texto cabe. */
export function applyEdit(current: NameDecision, patch: Partial<Pick<NameDecision, "nome" | "sobrenome">>): NameDecision {
  const next = { ...current, ...patch, edited: true };
  return { ...next, confirmed: fits(next.nome, next.sobrenome) };
}

export function isResolved(decision: NameDecision): boolean {
  return decision.confirmed && fits(decision.nome, decision.sobrenome);
}

export function pendingItems(items: NameReviewItem[], decisions: Decisions): NameReviewItem[] {
  return items.filter((item) => !isResolved(decisionFor(item, decisions)));
}

/** "Aceitar todas": confirma toda sugestão que cabe. As que não cabem (nem abreviadas) continuam pendentes: só
 *  editando à mão. */
export function acceptAll(items: NameReviewItem[], decisions: Decisions): Decisions {
  const next = { ...decisions };
  for (const item of items) {
    const decision = decisionFor(item, decisions);
    if (fits(decision.nome, decision.sobrenome)) next[item.key] = { ...decision, confirmed: true };
  }
  return next;
}

/** O que vai junto do arquivo na geração: a decisão final de cada nome conferido, por chave "arquivo:linha". */
export type NameOverrides = Record<string, { nome_completo: string; nome: string; sobrenome: string; editado: boolean }>;

export function buildOverrides(items: NameReviewItem[], decisions: Decisions): NameOverrides {
  const overrides: NameOverrides = {};
  for (const item of items) {
    const decision = decisionFor(item, decisions);
    if (!isResolved(decision)) continue;
    overrides[item.key] = {
      nome_completo: item.nome_completo,
      nome: sanitizeName(decision.nome),
      sobrenome: sanitizeName(decision.sobrenome),
      editado: decision.edited,
    };
  }
  return overrides;
}
