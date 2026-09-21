import type { QualityReport } from "./api";

function plural(n: number, singular: string, pluralForma: string): string {
  return `${n} ${n === 1 ? singular : pluralForma}`;
}

export interface Pendencias {
  /** O que invalida linhas (campo obrigatório em branco, CPF/e-mail malformado...) ou a ficha inteira:
   *  pede confirmação para gerar. */
  graves: string[];
  /** Só informam, sem ação do usuário: linhas repetidas o próprio sistema já tira do arquivo de carga. */
  avisos: string[];
}

/** O que a validação achou, em frases curtas ("3 linhas inválidas"). Única fonte para o resumo e para a
 *  confirmação antes de gerar — os dois dizem o mesmo. Tudo vazio = planilha limpa.
 *  Campo obrigatório em branco não aparece à parte: ele já invalida a linha, e o detalhe por campo
 *  fica no relatório (`required_blank`). */
export function pendenciasDoRelatorio(report: QualityReport): Pendencias {
  const graves: string[] = [];
  if (report.general_errors) graves.push("erro geral na planilha");
  if (report.invalid_rows > 0) graves.push(plural(report.invalid_rows, "linha inválida", "linhas inválidas"));

  const avisos: string[] = [];
  if (report.duplicated_rows > 0) {
    avisos.push(plural(report.duplicated_rows, "linha duplicada removida", "linhas duplicadas removidas"));
  }

  return { graves, avisos };
}
