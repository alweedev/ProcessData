import type { ReactNode } from "react";

interface EmptyStateProps {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
}

export function EmptyState({ icon, title, description, action }: EmptyStateProps) {
  return (
    <div className="flex flex-col items-center justify-center rounded-lg border border-dashed border-border px-6 py-10 text-center">
      {icon && (
        <div
          aria-hidden="true"
          className="mb-3 flex h-11 w-11 items-center justify-center rounded-full bg-surface-2 text-text-subtle"
        >
          {icon}
        </div>
      )}
      <p className="text-sm font-medium text-text">{title}</p>
      {description && <p className="mt-1 max-w-sm text-sm text-text-muted">{description}</p>}
      {action && <div className="mt-4">{action}</div>}
    </div>
  );
}
