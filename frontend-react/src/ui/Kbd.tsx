import type { ReactNode } from "react";

export function Kbd({ children }: { children: ReactNode }) {
  return (
    <kbd className="inline-flex min-w-[1.5rem] items-center justify-center rounded border border-border-strong bg-surface-2 px-1.5 py-0.5 font-mono text-xs text-text-muted">
      {children}
    </kbd>
  );
}
