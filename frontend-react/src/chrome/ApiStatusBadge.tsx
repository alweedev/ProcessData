import { checkApiHealth, useApiHealth, type ApiState } from "../health/apiHealthStore";

const TEXT: Record<ApiState, string> = {
  online: "API Online",
  offline: "API Offline",
  checking: "Verificando...",
};

function StatusIcon({ state }: { state: ApiState }) {
  if (state === "online") {
    return (
      <>
        <circle cx="12" cy="12" r="9" />
        <circle cx="12" cy="12" r="3" />
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          d="M12 9V5M12 19v-4M9 12H5M19 12h-4M9.6 9.6l-2.8-2.8M16.4 9.6l2.8-2.8M9.6 14.4l-2.8 2.8M16.4 14.4l2.8 2.8"
        />
      </>
    );
  }
  if (state === "checking") {
    return (
      <>
        <circle cx="12" cy="12" r="9" />
        <path strokeLinecap="round" strokeLinejoin="round" d="M12 3a9 9 0 019 9" />
        <circle cx="12" cy="12" r="2" />
      </>
    );
  }
  return (
    <>
      <circle cx="12" cy="12" r="9" />
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M8.5 13.5l3-3m1 1.5l2.2 2.2M14.5 10.5L16 9m-8 6l-1.5 1.5M9 8.8L7.5 7.3"
      />
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 6l12 12" />
    </>
  );
}

export function ApiStatusBadge() {
  const state = useApiHealth();

  const tone =
    state === "online"
      ? "border-success/40 text-success"
      : state === "offline"
        ? "border-danger/40 text-danger"
        : "border-border-strong text-text-subtle";

  return (
    <button
      id="apiStatusBtn"
      type="button"
      aria-label="Status da API"
      title="Status da API"
      onClick={() => void checkApiHealth(true)}
      className={`inline-flex items-center gap-1.5 rounded-lg border px-3 py-2 text-xs font-semibold uppercase ${tone}`}
    >
      <svg
        id="apiStatusIcon"
        xmlns="http://www.w3.org/2000/svg"
        className={`h-4 w-4${state === "checking" ? " animate-spin" : ""}`}
        fill="none"
        viewBox="0 0 24 24"
        stroke="currentColor"
        strokeWidth={1.8}
        aria-hidden="true"
      >
        <StatusIcon state={state} />
      </svg>
      <span id="apiStatusText" className="hidden sm:inline">{TEXT[state]}</span>
    </button>
  );
}
