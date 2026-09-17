import type { ReactNode } from "react";
import { cn } from "./cn";
import { IconChip } from "./IconChip";

interface EmptyStateProps {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
  /** Reduz o padding vertical — pra usos dentro de um card que já tem seu próprio respiro (ex.: Home). */
  compact?: boolean;
}

export function EmptyState({ icon, title, description, action, compact = false }: EmptyStateProps) {
  return (
    <div
      className={cn(
        "flex flex-col items-center justify-center rounded-surface border border-dashed border-border px-6 text-center",
        compact ? "py-8" : "py-10",
      )}
    >
      {icon && (
        <div className="mb-3">
          <IconChip icon={icon} size="lg" shape="circle" tone="neutral" />
        </div>
      )}
      <p className="text-sm font-medium text-text">{title}</p>
      {description && <p className="mt-1 max-w-sm text-sm text-text-muted">{description}</p>}
      {action && <div className="mt-4">{action}</div>}
    </div>
  );
}
