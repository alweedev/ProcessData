import { cn } from "./cn";
import { IconCheck } from "./icons";

export type StepState = "done" | "current" | "todo";

export interface StepperItem {
  label: string;
  /** Linha de apoio (ex.: "2 arquivos", "CPF · SELF"); some no mobile. */
  detail?: string;
  state: StepState;
}

const BUBBLE: Record<StepState, string> = {
  done: "border-success/40 bg-success-soft text-success",
  current: "border-accent bg-accent text-accent-fg",
  todo: "border-border-strong bg-surface text-text-subtle",
};

const STATE_HINT: Record<StepState, string> = {
  done: " (concluída)",
  current: " (etapa atual)",
  todo: "",
};

/** Indicador de progresso do fluxo (só informativo — as seções continuam
 *  todas acessíveis). Horizontal, com conectores entre as etapas. */
export function Stepper({ steps, label }: { steps: StepperItem[]; label: string }) {
  return (
    <ol aria-label={label} className="mb-5 flex items-start">
      {steps.map((step, index) => (
        <li
          key={step.label}
          aria-current={step.state === "current" ? "step" : undefined}
          className={cn("flex items-start", index < steps.length - 1 ? "flex-1" : "flex-none")}
        >
          <div className="flex min-w-0 items-center gap-2 sm:gap-3">
            <span
              aria-hidden="true"
              className={cn(
                "flex h-8 w-8 shrink-0 items-center justify-center rounded-pill border text-sm font-semibold transition-colors",
                BUBBLE[step.state],
              )}
            >
              {step.state === "done" ? <IconCheck className="h-4 w-4" /> : index + 1}
            </span>
            {/* Mobile: só a etapa atual mostra o rótulo; as demais ficam só no círculo. */}
            <div className={cn("min-w-0", step.state !== "current" && "sr-only sm:not-sr-only")}>
              <p className={cn("truncate text-sm font-medium", step.state === "todo" ? "text-text-muted" : "text-text")}>
                {step.label}
                <span className="sr-only">{STATE_HINT[step.state]}</span>
              </p>
              {step.detail && <p className="hidden truncate text-xs text-text-subtle sm:block">{step.detail}</p>}
            </div>
          </div>
          {index < steps.length - 1 && (
            <span
              aria-hidden="true"
              className={cn(
                "mx-2 mt-4 h-px min-w-4 flex-1 transition-colors sm:mx-3 sm:min-w-6",
                step.state === "done" ? "bg-success/40" : "bg-border",
              )}
            />
          )}
        </li>
      ))}
    </ol>
  );
}
