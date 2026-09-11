type Tone = "neutral" | "success" | "warning" | "danger" | "info";

interface StatCardProps {
  label: string;
  value: string | number;
  tone?: Tone;
  hint?: string;
}

const TONE: Record<Tone, string> = {
  neutral: "text-text",
  success: "text-success",
  warning: "text-warning",
  danger: "text-danger",
  info: "text-info",
};

export function StatCard({ label, value, tone = "neutral", hint }: StatCardProps) {
  return (
    <div className="rounded-lg border border-border bg-surface px-3 py-2.5">
      <div className="text-xs font-medium uppercase tracking-wide text-text-subtle">{label}</div>
      <div className={`mt-1 text-2xl font-semibold tabular-nums ${TONE[tone]}`}>{value}</div>
      {hint && <div className="mt-0.5 text-xs text-text-muted">{hint}</div>}
    </div>
  );
}
