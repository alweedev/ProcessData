import { useState } from "react";
import type { QualityReport } from "../lib/api";
import { coveredByGeneral, generalProblems, lineDetails, summarizeProblems, type ProblemRow } from "../lib/problems";
import { cn } from "./cn";

interface ValidationReportProps {
  report: QualityReport | null;
  error?: string | null;
}

const VISIBLE_ROWS = 5;

/**
 * O detalhe das pendências, para quem vai corrigir a planilha. Vai DENTRO do cartão do ValidationSummary (que já diz
 * quantos cadastros têm pendência), então não tem título nem contagem própria e não repete nada:
 * - problema da ficha inteira (coluna ausente): uma frase, e não o mesmo "em branco" de cada passageiro;
 * - o que falta em todas as linhas, uma vez; cada passageiro uma vez, só com o que ele tem a mais;
 * - linhas sem nenhum obrigatório: "parece que não foram preenchidas", sem listar os problemas de cada uma.
 * Sem problema algum não mostra nada. O erro de validação (servidor recusou, rede) aparece sozinho, sem cartão.
 */
export function ValidationReport({ report, error }: ValidationReportProps) {
  if (error) {
    return (
      <div className="mb-4 rounded-surface border border-border bg-surface-2 px-4 py-4 text-sm text-text-muted">
        {error}
      </div>
    );
  }
  if (!report) return null;

  const general = generalProblems(report);
  const { blankRows, common, rows } = summarizeProblems(lineDetails(report), coveredByGeneral(report));
  if (general.length === 0 && blankRows.length === 0 && rows.length === 0) return null;

  return (
    <div id="cadastro_problems" className="space-y-3">
      {general.length > 0 && (
        <ul className="space-y-1.5">
          {general.map((problem) => (
            <li key={problem} className="rounded-control border border-danger/40 px-3 py-2 text-sm text-danger">
              {problem}
            </li>
          ))}
        </ul>
      )}

      {blankRows.length > 0 && (
        <div id="cadastro_problems_blank">
          <p className="text-sm font-medium text-text">
            {blankRows.length}{" "}
            {blankRows.length === 1 ? "linha sem nenhum campo obrigatório" : "linhas sem nenhum campo obrigatório"}
            <span className="font-normal text-text-muted">
              {blankRows.length === 1 ? " — parece que não foi preenchida" : " — parece que não foram preenchidas"}
            </span>
          </p>
          <RowList rows={blankRows.map((row) => ({ ...row, extras: [] }))} />
        </div>
      )}

      {rows.length > 0 && (
        <div id="cadastro_problems_rows" className={blankRows.length > 0 ? "border-t border-border pt-3" : undefined}>
          {common.length > 0 && (
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="text-sm font-medium text-text">
                {rows.length === 2 ? "Nos dois:" : `Em todos os ${rows.length}:`}
              </span>
              <ul className="contents">
                {common.map((title) => (
                  <li key={title} className="rounded-full border border-danger/40 px-2 py-0.5 text-xs text-danger">
                    {title}
                  </li>
                ))}
              </ul>
            </div>
          )}
          <RowList rows={rows} alsoPrefix={common.length > 0} />
        </div>
      )}
    </div>
  );
}

type ListedRow = ProblemRow & { extras: string[] };

function RowList({ rows, alsoPrefix = false }: { rows: ListedRow[]; alsoPrefix?: boolean }) {
  const [expanded, setExpanded] = useState(false);
  const visible = expanded ? rows : rows.slice(0, VISIBLE_ROWS);
  const hidden = rows.length - visible.length;

  return (
    <>
      <ul className={cn("mt-1.5 space-y-0.5 text-xs text-text-muted", expanded && "max-h-64 overflow-y-auto")}>
        {visible.map((row, index) => (
          <li key={`${row.label}-${index}`}>
            <span className="font-medium text-text-subtle">{row.label}</span>
            {row.nome && (
              <>
                {" · "}
                <span className="text-text">{row.nome}</span>
              </>
            )}
            {row.extras.length > 0 && (
              <span>
                {" — "}
                {alsoPrefix ? "também: " : ""}
                {row.extras.join("; ")}
              </span>
            )}
          </li>
        ))}
      </ul>
      {rows.length > VISIBLE_ROWS && (
        <button
          type="button"
          aria-expanded={expanded}
          onClick={() => setExpanded((v) => !v)}
          className="mt-1.5 rounded-control text-xs font-medium text-accent-text outline-none hover:underline focus-visible:ring-2 focus-visible:ring-accent/50"
        >
          {expanded ? "Mostrar menos" : `Mostrar todos (${hidden} a mais)`}
        </button>
      )}
    </>
  );
}
