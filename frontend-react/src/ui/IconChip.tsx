import type { ReactNode } from "react";
import { cn } from "./cn";
import { TONE_SOFT, type Tone } from "./tone";

type Size = "sm" | "md" | "lg";
type Shape = "circle" | "square";
type IconChipTone = Extract<Tone, "brand" | "neutral" | "success" | "danger">;

const SIZE: Record<Size, string> = { sm: "h-7 w-7", md: "h-9 w-9", lg: "h-11 w-11" };
const SHAPE: Record<Shape, string> = { circle: "rounded-pill", square: "rounded-control" };

interface IconChipProps {
  icon: ReactNode;
  size?: Size;
  shape?: Shape;
  tone?: IconChipTone;
}

/** Selo decorativo com ícone — unifica o padrão antes duplicado (e com
 *  tamanhos/formas divergentes) entre PageHeader, cards da Home, EmptyState
 *  e ResultCard. Reusa TONE_SOFT (tone.ts) em vez de manter seu próprio mapa
 *  de cores, pra não divergir do resto do app (ex.: Badge). */
export function IconChip({ icon, size = "md", shape = "square", tone = "brand" }: IconChipProps) {
  return (
    <span
      aria-hidden="true"
      className={cn("flex shrink-0 items-center justify-center", SIZE[size], SHAPE[shape], TONE_SOFT[tone])}
    >
      {icon}
    </span>
  );
}
