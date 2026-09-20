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

/** Estado de cada ponto: o que está sendo visto é "current"; o resto, concluído se já preenchido. */
export function estadosDaLinhaDoTempo(atual: number, preenchido: Preenchido, concluido: boolean): StepState[] {
  return [0, 1, 2, 3].map((etapa): StepState => {
    if (concluido) return "done";
    if (etapa === atual) return "current";
    return etapa < ULTIMA_ETAPA && preenchido[etapa] ? "done" : "todo";
  });
}
