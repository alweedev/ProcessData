import { TONE_TEXT, type Tone } from "./tone";

type StatTone = Exclude<Tone, "brand">;

interface StatCardProps {
  label: string;
  value: string | number;
  tone?: StatTone;
  hint?: string;
}

export function StatCard({ label, value, tone = "neutral", hint }: StatCardProps) {
  return (
    <div className="rounded-surface border border-border bg-surface px-3 py-2.5">
      <div className="text-xs font-medium uppercase tracking-wide text-text-subtle">{label}</div>
      <div className={`mt-1 text-2xl font-semibold tabular-nums ${TONE_TEXT[tone]}`}>{value}</div>
      {hint && <div className="mt-0.5 text-xs text-text-muted">{hint}</div>}
    </div>
  );
}
