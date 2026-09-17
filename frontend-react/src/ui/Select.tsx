import type { SelectHTMLAttributes } from "react";
import { cn } from "./cn";

export function Select({ className, children, ...rest }: SelectHTMLAttributes<HTMLSelectElement>) {
  return (
    <select
      className={cn(
        "w-full rounded-control border border-border bg-surface px-3 py-2 text-sm text-text outline-none transition-colors focus:border-accent focus:ring-2 focus:ring-accent/30",
        className,
      )}
      {...rest}
    >
      {children}
    </select>
  );
}
