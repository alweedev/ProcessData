import type { AnaliseInativacao } from "../../lib/inativacaoApi";

export interface Confirmacao {
  impacto: boolean;
  orfas: boolean;
}

/** Do Impacto para o Confirmar: há quem inativar e nenhum homônimo ficou sem escolha. */
export function podeContinuar(analise: AnaliseInativacao | null): boolean {
  if (!analise) return false;
  return analise.resumo.executaveis > 0 && !analise.usuarios.some((u) => u.situacao === "PENDENTE_SELECAO");
}

/** Executar exige a confirmação do impacto e, havendo estrutura órfã, também a ciência dela. */
export function podeExecutar(analise: AnaliseInativacao | null, confirmacao: Confirmacao): boolean {
  if (!analise || !podeContinuar(analise)) return false;
  return confirmacao.impacto && (analise.resumo.estruturasOrfas === 0 || confirmacao.orfas);
}
