import type { ReactNode } from "react";
import { IconChip } from "./IconChip";

interface ResultCardProps {
  operationLabel: string;
  timestamp: number;
  inputSummary: string | string[];
  outputFilename?: string;
  status?: "success" | "error";
  onRedownload?: () => void;
  compact?: boolean;
  icon?: ReactNode;
}

function fmtTime(ts: number): string {
  try {
    return new Date(ts).toLocaleString("pt-BR");
  } catch {
    return "";
  }
}

export function ResultCard({
  operationLabel,
  timestamp,
  inputSummary,
  outputFilename,
  status = "success",
  onRedownload,
  compact = false,
  icon,
}: ResultCardProps) {
  const lines = Array.isArray(inputSummary) ? inputSummary : [inputSummary];
  return (
    <div
      className={`flex items-start gap-3 rounded-surface border border-border bg-surface ${compact ? "px-3 py-2" : "p-3.5"}`}
    >
      <span className="mt-0.5 text-sm">
        <IconChip
          icon={icon ?? (status === "success" ? "✓" : "!")}
          size="sm"
          shape="circle"
          tone={status === "success" ? "success" : "danger"}
        />
      </span>
      <div className="min-w-0 flex-1">
        <div className="flex flex-wrap items-baseline justify-between gap-x-3 gap-y-0.5">
          <span className="text-sm font-medium text-text">{operationLabel}</span>
          <span className="text-xs text-text-subtle">{fmtTime(timestamp)}</span>
        </div>
        <div className="mt-0.5 space-y-0.5">
          {lines.map((line, i) => (
            <p key={i} className="truncate text-xs text-text-muted">
              {line}
            </p>
          ))}
        </div>
        {(outputFilename || onRedownload) && (
          <div className="mt-1.5 flex flex-wrap items-center gap-2">
            {outputFilename && <span className="font-mono text-xs text-text-subtle">{outputFilename}</span>}
            {onRedownload && (
              <button type="button" onClick={onRedownload} className="text-xs font-medium text-accent hover:underline">
                Baixar novamente
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
