import type { ReactNode } from "react";
import { TONE_OUTLINE, type Tone } from "./tone";

type Size = "sm" | "md";

interface BadgeProps {
  tone?: Tone;
  size?: Size;
  children: ReactNode;
}

export function Badge({ tone = "neutral", size = "sm", children }: BadgeProps) {
  return (
    <span
      className={`inline-flex items-center gap-1 rounded-pill border font-medium ${TONE_OUTLINE[tone]} ${
        size === "sm" ? "px-2 py-0.5 text-xs" : "px-2.5 py-1 text-sm"
      }`}
    >
      {children}
    </span>
  );
}
