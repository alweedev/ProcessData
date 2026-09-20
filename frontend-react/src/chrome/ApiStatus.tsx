import { useApiHealth, type ApiState } from "../health/apiHealthStore";
import { Badge } from "../ui/Badge";
import type { Tone } from "../ui/tone";

const TEXT: Record<ApiState, string> = {
  online: "API online",
  offline: "API offline",
  checking: "Verificando…",
};
const TONE: Record<ApiState, Tone> = {
  online: "success",
  offline: "danger",
  checking: "neutral",
};
const DOT: Record<ApiState, string> = {
  online: "bg-success",
  offline: "bg-danger",
  checking: "bg-text-subtle animate-pulse",
};

export function ApiStatus() {
  const state = useApiHealth();
  return (
    <span role="status" aria-live="polite">
      <Badge tone={TONE[state]}>
        <span aria-hidden="true" className={`h-1.5 w-1.5 rounded-full ${DOT[state]}`} />
        {TEXT[state]}
      </Badge>
    </span>
  );
}
