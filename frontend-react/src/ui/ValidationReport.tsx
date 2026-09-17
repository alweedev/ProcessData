import type { QualityReport } from "../lib/api";
import { StatCard } from "./StatCard";
import { TONE_OUTLINE } from "./tone";

interface ValidationReportProps {
  report: QualityReport | null;
  loading: boolean;
  error?: string | null;
}

export function ValidationReport({ report, loading, error }: ValidationReportProps) {
  if (loading) {
    return (
      <div className="rounded-surface border border-border bg-surface-2 px-4 py-6 text-center text-sm text-text-muted">
        Validando planilha…
      </div>
    );
  }
  if (error) {
    return (
      <div className="rounded-surface border border-border bg-surface-2 px-4 py-4 text-sm text-text-muted">{error}</div>
    );
  }
  if (!report) return null;

  const lineErrors = Object.entries(report.line_errors);
  const blanks = Object.entries(report.required_blank).filter(([, n]) => n > 0);
  const clean =
    report.invalid_rows === 0 && report.duplicated_rows === 0 && !report.general_errors && blanks.length === 0;

  return (
    <div className="space-y-3 rounded-surface border border-border bg-surface-2 p-4">
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
        <StatCard label="Linhas" value={report.total_rows} />
        <StatCard label="Válidas" value={report.valid_rows} tone={report.valid_rows > 0 ? "success" : "neutral"} />
        <StatCard label="Inválidas" value={report.invalid_rows} tone={report.invalid_rows > 0 ? "danger" : "neutral"} />
        <StatCard
          label="Duplicadas"
          value={report.duplicated_rows}
          tone={report.duplicated_rows > 0 ? "warning" : "neutral"}
        />
      </div>

      {clean && <p className="text-sm text-success">Nenhum problema encontrado — pode gerar.</p>}

      {report.general_errors && (
        <div className={`rounded-control border px-3 py-2 text-sm ${TONE_OUTLINE.danger}`}>{report.general_errors}</div>
      )}

      {blanks.length > 0 && (
        <div className="text-sm text-text">
          <p className="mb-1 font-medium">Campos obrigatórios em branco</p>
          <div className="flex flex-wrap gap-1.5">
            {blanks.map(([col, n]) => (
              <span key={col} className={`rounded-pill border px-2 py-0.5 text-xs ${TONE_OUTLINE.warning}`}>
                {col}: {n}
              </span>
            ))}
          </div>
        </div>
      )}

      {lineErrors.length > 0 && (
        <div className="text-sm text-text">
          <p className="mb-1 font-medium">Erros por linha ({lineErrors.length})</p>
          <ul className="max-h-48 space-y-1 overflow-y-auto rounded-control bg-surface px-3 py-2 text-xs text-text-muted">
            {lineErrors.map(([line, msg]) => (
              <li key={line}>
                <span className="font-mono text-text-subtle">L{line}</span> — {msg}
              </li>
            ))}
          </ul>
        </div>
      )}

      <p className="text-xs text-text-subtle">A validação é informativa — a geração não depende dela.</p>
    </div>
  );
}
