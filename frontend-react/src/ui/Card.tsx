import type { HTMLAttributes, ReactNode } from "react";
import { cn } from "./cn";

interface CardProps extends HTMLAttributes<HTMLDivElement> {
  interactive?: boolean;
  padding?: "none" | "sm" | "md" | "lg";
  children: ReactNode;
}

const PAD: Record<NonNullable<CardProps["padding"]>, string> = {
  none: "",
  sm: "p-3",
  md: "p-4 sm:p-5",
  lg: "p-5 sm:p-6",
};

export function Card({ interactive = false, padding = "md", className, children, ...rest }: CardProps) {
  return (
    <div
      className={cn(
        "rounded-surface border border-border bg-surface shadow-card",
        PAD[padding],
        interactive &&
          "cursor-pointer transition hover:-translate-y-1 hover:border-border-strong hover:shadow-card-hover",
        className,
      )}
      {...rest}
    >
      {children}
    </div>
  );
}
