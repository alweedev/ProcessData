import { useId, type KeyboardEvent } from "react";
import { cn } from "./cn";

export interface SegmentedOption {
  value: string;
  label: string;
  /** Explicação curta exibida sob o rótulo. */
  description?: string;
}

interface SegmentedControlProps {
  id: string;
  label: string;
  value: string;
  options: SegmentedOption[];
  onChange: (value: string) => void;
}

const ARROW_STEP: Record<string, number> = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 };

/** Escolha exclusiva entre poucas opções (radiogroup): mostra todas de uma
 *  vez, no lugar de um <select> que as esconde. Setas movem a seleção; só a
 *  opção marcada entra na ordem de Tab. */
export function SegmentedControl({ id, label, value, options, onChange }: SegmentedControlProps) {
  const labelId = useId();

  function onKeyDown(e: KeyboardEvent<HTMLButtonElement>, index: number) {
    const step = ARROW_STEP[e.key];
    if (!step) return;
    e.preventDefault();
    const next = options[(index + step + options.length) % options.length];
    onChange(next.value);
    document.getElementById(`${id}-${next.value}`)?.focus();
  }

  return (
    <div>
      <p id={labelId} className="mb-1.5 text-sm font-medium text-text">
        {label}
      </p>
      <div id={id} role="radiogroup" aria-labelledby={labelId} className="grid grid-cols-2 gap-2">
        {options.map((option, index) => {
          const checked = option.value === value;
          return (
            <button
              key={option.value}
              id={`${id}-${option.value}`}
              type="button"
              role="radio"
              aria-checked={checked}
              tabIndex={checked ? 0 : -1}
              onClick={() => onChange(option.value)}
              onKeyDown={(e) => onKeyDown(e, index)}
              className={cn(
                "rounded-control border px-3 py-2 text-left outline-none transition-colors focus-visible:ring-2 focus-visible:ring-accent/50",
                checked ? "border-accent bg-accent/10" : "border-border bg-surface hover:border-border-strong hover:bg-surface-2",
              )}
            >
              <span className={cn("block text-sm font-semibold", checked ? "text-accent" : "text-text")}>
                {option.label}
              </span>
              {option.description && <span className="mt-0.5 block text-xs text-text-muted">{option.description}</span>}
            </button>
          );
        })}
      </div>
    </div>
  );
}
