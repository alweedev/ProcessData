import type { ReactNode } from "react";

interface StepSectionProps {
  number: number;
  title: string;
  description?: string;
  children: ReactNode;
}

/** Bloco numerado de um fluxo em etapas: dentro de um mesmo Card, separado
 *  do anterior por uma divisória. */
export function StepSection({ number, title, description, children }: StepSectionProps) {
  return (
    <section className="border-t border-border py-5 first:border-t-0 first:pt-0 last:pb-0">
      <header className="mb-3 flex items-baseline gap-2">
        <span aria-hidden="true" className="text-xs font-semibold tabular-nums text-brand">
          {String(number).padStart(2, "0")}
        </span>
        <h3 className="text-base font-semibold text-text">{title}</h3>
        {description && <p className="hidden text-sm text-text-muted sm:block">— {description}</p>}
      </header>
      {children}
    </section>
  );
}
