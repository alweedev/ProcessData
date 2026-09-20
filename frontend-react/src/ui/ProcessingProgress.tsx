import { cn } from "./cn";

interface ProcessingProgressProps {
  /** Id da trilha; a barra recebe `${id}Bar` (ids usados pelos testes e2e). */
  id: string;
  /** Porcentagem do ENVIO (0–100), como reportada por postFormForBlob. */
  progress: number;
}

/**
 * Progresso em duas fases. O percentual só mede o envio do arquivo; quando ele
 * chega a 100% o servidor ainda está processando e não há como saber quanto
 * falta, então a barra vira indeterminada em vez de ficar parada em "100%".
 */
export function ProcessingProgress({ id, progress }: ProcessingProgressProps) {
  const uploading = progress < 100;
  const pct = Math.round(progress);

  return (
    <div>
      <div className="mb-1 flex items-center justify-between text-xs text-text-muted" aria-live="polite">
        <span>{uploading ? "Enviando fichas…" : "Processando no servidor…"}</span>
        {uploading && <span className="tabular-nums">{pct}%</span>}
      </div>
      <div id={id} className="h-2 overflow-hidden rounded-pill bg-surface-sunken">
        <div
          id={`${id}Bar`}
          role="progressbar"
          aria-label={uploading ? "Envio das fichas" : "Processamento no servidor"}
          aria-valuemin={0}
          aria-valuemax={100}
          aria-valuenow={uploading ? pct : undefined}
          style={uploading ? { width: `${pct}%` } : undefined}
          className={cn(
            "h-full bg-accent",
            uploading
              ? "transition-[width]"
              : "w-1/4 animate-[pd-indeterminate_1.4s_ease-in-out_infinite] motion-reduce:w-full motion-reduce:animate-none motion-reduce:opacity-60",
          )}
        />
      </div>
    </div>
  );
}
