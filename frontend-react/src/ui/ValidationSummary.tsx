import { useId, useState, type ReactNode } from "react";
import type { QualityReport } from "../lib/api";
import { pendenciasDoRelatorio } from "../lib/qualityReport";
import { cn } from "./cn";
import { IconAlert, IconCheck } from "./icons";
import { StatCard } from "./StatCard";
import { TONE_OUTLINE } from "./tone";

interface ValidationSummaryProps {
  status: "idle" | "loading" | "done" | "error";
  report: QualityReport | null;
  /** O detalhe das pendências (ValidationReport): vai dentro do mesmo cartão, só quando há pendência. */
  children?: ReactNode;
}

function plural(n: number, singular: string, pluralForma: string): string {
  return n === 1 ? singular : pluralForma;
}

/**
 * O veredito da validação, para quem só quer saber se pode gerar:
 * - tudo certo: UMA linha ("Tudo certo: 5 cadastros prontos para gerar"), com os números atrás de "Ver detalhes";
 * - só aviso (linha repetida removida): a mesma linha, avisando o que foi feito;
 * - pendências: um cartão só com quantos cadastros têm problema, o que fazer e, dentro dele, o detalhe (`children`,
 *   o ValidationReport): nada é dito duas vezes em cartões diferentes.
 * Erro de validação não aparece aqui: o relatório já o mostra e a geração não depende dele.
 */
export function ValidationSummary({ status, report, children }: ValidationSummaryProps) {
  const [open, setOpen] = useState(false);
  const detailsId = useId();

  if (status === "loading") {
    return (
      <div
        role="status"
        id="cadastro_validation_summary"
        className="mb-4 flex items-center gap-3 rounded-control border border-border bg-surface-2 px-3 py-2.5 text-sm text-text-muted"
      >
        <svg className="h-4 w-4 shrink-0 animate-spin" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z" />
        </svg>
        Validando a planilha…
      </div>
    );
  }
  if (status !== "done" || !report) return null;

  const { graves, avisos } = pendenciasDoRelatorio(report);
  const tom = graves.length > 0 ? "warning" : avisos.length > 0 ? "info" : "success";
  const prontos = `${report.valid_rows} ${plural(report.valid_rows, "cadastro pronto", "cadastros prontos")} para gerar`;
  const withDetails = tom !== "warning";

  return (
    <div
      role="status"
      id="cadastro_validation_summary"
      data-tone={tom}
      className={cn("mb-4 rounded-control border px-3 py-2.5 text-sm", TONE_OUTLINE[tom])}
    >
      <div className="flex flex-wrap items-start gap-x-3 gap-y-1">
        <span className="mt-0.5 shrink-0">
          {tom === "warning" ? <IconAlert className="h-4 w-4" /> : <IconCheck className="h-4 w-4" />}
        </span>
        <div className="min-w-0 flex-1 text-text">
          {tom === "success" && (
            <p>
              <strong className="font-medium text-success">Tudo certo:</strong> {prontos}.
            </p>
          )}
          {tom === "info" && (
            <p>
              <strong className="font-medium text-info">Tudo certo, com um aviso:</strong> {avisos.join(" · ")}.{" "}
              {prontos}.
            </p>
          )}
          {tom === "warning" && <Pendencias report={report} />}
        </div>
        {withDetails && (
          <button
            type="button"
            aria-expanded={open}
            aria-controls={detailsId}
            onClick={() => setOpen((v) => !v)}
            className="shrink-0 rounded-control px-1.5 py-0.5 text-xs font-medium text-text-muted underline-offset-2 outline-none hover:text-text hover:underline focus-visible:ring-2 focus-visible:ring-accent/50"
          >
            {open ? "Ocultar detalhes" : "Ver detalhes"}
          </button>
        )}
      </div>

      {tom === "warning" && children && <div className="mt-3 border-t border-border pt-3 text-text">{children}</div>}

      {withDetails && open && (
        <div id={detailsId} className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-4">
          <StatCard label="Linhas" value={report.total_rows} />
          <StatCard label="Válidas" value={report.valid_rows} tone={report.valid_rows > 0 ? "success" : "neutral"} />
          <StatCard
            label="Inválidas"
            value={report.invalid_rows}
            tone={report.invalid_rows > 0 ? "danger" : "neutral"}
          />
          <StatCard
            label="Duplicadas"
            value={report.duplicated_rows}
            tone={report.duplicated_rows > 0 ? "warning" : "neutral"}
          />
        </div>
      )}
    </div>
  );
}

/** Quantos cadastros têm pendência e o que fazer. Só conta o que não está no título: prontos e repetidas removidas. */
function Pendencias({ report }: { report: QualityReport }) {
  const { total_rows: total, valid_rows: valid, invalid_rows: invalid, duplicated_rows: repetidas } = report;
  const headline =
    invalid === 0
      ? "A planilha tem um problema geral."
      : total === 1
        ? "O cadastro tem pendência."
        : `${invalid} de ${total} cadastros ${invalid === 1 ? "tem" : "têm"} pendência.`;
  const contas = [
    valid > 0 && invalid > 0 ? `${valid} ${plural(valid, "pronto", "prontos")}` : "",
    repetidas > 0
      ? `${repetidas} ${plural(repetidas, "linha repetida removida", "linhas repetidas removidas")} (a planilha tinha ${total + repetidas})`
      : "",
  ].filter(Boolean);

  return (
    <>
      <p>
        <strong className="font-medium text-warning">{headline}</strong>
        {contas.length > 0 && <span className="text-text-muted"> {contas.join(" · ")}</span>}
      </p>
      <p className="mt-0.5 text-text-muted">Corrija a ficha ou gere assim mesmo (pediremos sua confirmação).</p>
    </>
  );
}
