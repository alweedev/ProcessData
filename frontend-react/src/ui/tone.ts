export type Tone = "neutral" | "brand" | "success" | "warning" | "danger" | "info";

/** Fundo suave + texto colorido — pills de status (Badge, avisos inline). */
export const TONE_SOFT: Record<Tone, string> = {
  neutral: "bg-surface-sunken text-text-muted",
  brand: "bg-brand-soft text-brand",
  success: "bg-success-soft text-success",
  warning: "bg-warning-soft text-warning",
  danger: "bg-danger-soft text-danger",
  info: "bg-info-soft text-info",
};

/** Só cor de texto — valores numéricos/destacados (StatCard). */
export const TONE_TEXT: Record<Tone, string> = {
  neutral: "text-text",
  brand: "text-brand",
  success: "text-success",
  warning: "text-warning",
  danger: "text-danger",
  info: "text-info",
};

/** Contornado, sem fundo — alertas persistentes (ex.: ApiHealthBanner). */
export const TONE_OUTLINE: Record<Tone, string> = {
  neutral: "border-border-strong text-text-muted",
  brand: "border-brand/40 text-brand",
  success: "border-success/40 text-success",
  warning: "border-warning/40 text-warning",
  danger: "border-danger/40 text-danger",
  info: "border-info/40 text-info",
};
