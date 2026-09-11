import type { TextareaHTMLAttributes } from "react";
import { cn } from "./cn";

export function Textarea({ className, ...rest }: TextareaHTMLAttributes<HTMLTextAreaElement>) {
  return (
    <textarea
      className={cn(
        "w-full rounded-lg border border-border bg-surface px-3 py-2 text-sm text-text outline-none transition-colors placeholder:text-text-subtle focus:border-accent focus:ring-2 focus:ring-accent/30",
        className,
      )}
      {...rest}
    />
  );
}
