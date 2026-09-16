import type { PreviewItem } from "./useEstruturas";

export function computeContexto(item: PreviewItem): string {
  const cod = item.ccCodigo || "";
  const desc = item.ccDescricao || "";
  if (item.aprovacaoPor === "VIAJANTE" && item.viajanteNomeCompleto) return item.viajanteNomeCompleto;
  if (item.aprovacaoPor === "CCEMPRESA") {
    if (cod && desc) return `${cod} - ${desc}`;
    return cod || desc || "";
  }
  if (cod && desc) return `${cod} - ${desc}`;
  if (item.viajanteNomeCompleto) return item.viajanteNomeCompleto;
  return cod || desc || "";
}
