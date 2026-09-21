import type { StepState } from "../../ui/Stepper";

/** Etapas do Cadastro: 0 fichas, 1 tipo de login, 2 fluxo, 3 gerar. */
export const ULTIMA_ETAPA = 3;

/** O que já está preenchido nas etapas 0..2, na ordem: fichas, login, fluxo. */
export type Preenchido = readonly [boolean, boolean, boolean];

/** Até onde dá para ir: a primeira etapa ainda vazia, ou a última se tudo está preenchido. */
export function etapaMaxima(preenchido: Preenchido): number {
  const primeiraVazia = preenchido.findIndex((ok) => !ok);
  return primeiraVazia === -1 ? ULTIMA_ETAPA : primeiraVazia;
}

/** Um ponto da linha do tempo é clicável se está ao alcance; depois de concluir, nenhum é. */
export function podeIrPara(etapa: number, preenchido: Preenchido, concluido: boolean): boolean {
  return !concluido && etapa <= etapaMaxima(preenchido);
}

/**
 * Estado de cada ponto. A linha do tempo mostra o progresso, não só o que está preenchido:
 * - a etapa que está sendo vista é "current";
 * - "done" só para as etapas ANTES dela que estão de fato concluídas, ou seja, preenchidas e
 *   com tudo o que vem antes também preenchido (sem fichas, login e fluxo não valem nada);
 * - "ready" para as etapas DEPOIS dela com valor guardado nas mesmas condições: voltar apaga o
 *   verde adiante, mas o ponto deixa claro que a escolha continua lá (e que dá para pular direto);
 * - o resto é "todo".
 */
export function estadosDaLinhaDoTempo(atual: number, preenchido: Preenchido, concluido: boolean): StepState[] {
  const alcance = etapaMaxima(preenchido);
  return [0, 1, 2, 3].map((etapa): StepState => {
    if (concluido) return "done";
    if (etapa === atual) return "current";
    if (etapa >= alcance) return "todo";
    return etapa < atual ? "done" : "ready";
  });
}
