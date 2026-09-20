import { useId, type KeyboardEvent } from "react";
import { cn } from "./cn";

export interface SegmentedOption {
  value: string;
  label: string;
}

interface SegmentedControlProps {
  id: string;
  label: string;
  /** `null` = nenhuma opção escolhida (não há valor padrão). */
  value: string | null;
  options: SegmentedOption[];
  onChange: (value: string) => void;
  /** Escolha obrigatória: anuncia `aria-required` e mostra "obrigatório" enquanto não houver escolha. */
  required?: boolean;
}

const ARROW_STEP: Record<string, number> = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 };

/** Escolha exclusiva entre poucas opções (radiogroup): mostra todas de uma
 *  vez, no lugar de um <select> que as esconde. Setas movem a seleção; só a
 *  opção marcada entra na ordem de Tab. */
export function SegmentedControl({ id, label, value, options, onChange, required = false }: SegmentedControlProps) {
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
      <div className="mb-1.5 flex items-baseline gap-2">
        <p id={labelId} className="text-sm font-medium text-text">
          {label}
        </p>
        {required && value === null && <span className="text-xs text-text-muted">obrigatório</span>}
      </div>
      <div
        id={id}
        role="radiogroup"
        aria-labelledby={labelId}
        aria-required={required || undefined}
        className="grid grid-cols-2 gap-2"
      >
        {options.map((option, index) => {
          const checked = option.value === value;
          return (
            <button
              key={option.value}
              id={`${id}-${option.value}`}
              type="button"
              role="radio"
              aria-checked={checked}
              // Sem escolha, a 1ª opção segue alcançável por Tab (senão o grupo ficaria inacessível).
              tabIndex={checked || (value === null && index === 0) ? 0 : -1}
              onClick={() => onChange(option.value)}
              onKeyDown={(e) => onKeyDown(e, index)}
              className={cn(
                "rounded-control border px-3 py-2.5 text-center outline-none transition-colors focus-visible:ring-2 focus-visible:ring-accent/50",
                checked ? "border-accent bg-accent/10" : "border-border bg-surface hover:border-border-strong hover:bg-surface-2",
              )}
            >
              <span className={cn("block text-sm font-semibold", checked ? "text-accent-text" : "text-text")}>
                {option.label}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
