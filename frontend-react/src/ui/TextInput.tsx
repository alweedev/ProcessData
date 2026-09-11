import type { InputHTMLAttributes } from "react";
import { cn } from "./cn";

export function TextInput({ className, type = "text", ...rest }: InputHTMLAttributes<HTMLInputElement>) {
  return (
    <input
      type={type}
      className={cn(
        "w-full rounded-lg border border-border bg-surface px-3 py-2 text-sm text-text outline-none transition-colors placeholder:text-text-subtle focus:border-accent focus:ring-2 focus:ring-accent/30",
        className,
      )}
      {...rest}
    />
  );
}
