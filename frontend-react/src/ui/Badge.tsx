import type { ReactNode } from "react";

type Tone = "neutral" | "brand" | "success" | "warning" | "danger" | "info";
type Size = "sm" | "md";

interface BadgeProps {
  tone?: Tone;
  size?: Size;
  children: ReactNode;
}

const TONE: Record<Tone, string> = {
  neutral: "bg-surface-sunken text-text-muted",
  brand: "bg-brand/12 text-brand",
  success: "bg-success-soft text-success",
  warning: "bg-warning-soft text-warning",
  danger: "bg-danger-soft text-danger",
  info: "bg-info-soft text-info",
};

export function Badge({ tone = "neutral", size = "sm", children }: BadgeProps) {
  return (
    <span
      className={`inline-flex items-center gap-1 rounded-full font-medium ${TONE[tone]} ${
        size === "sm" ? "px-2 py-0.5 text-xs" : "px-2.5 py-1 text-sm"
      }`}
    >
      {children}
    </span>
  );
}
