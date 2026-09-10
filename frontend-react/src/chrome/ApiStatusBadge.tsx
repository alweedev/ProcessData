import { useCallback, useEffect, useRef, useState } from "react";
import { pushToast } from "../toast/toastStore";

type ApiState = "online" | "offline" | "checking";

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

async function checkHealth(): Promise<boolean> {
  for (const url of ["/api/health", "/"]) {
    try {
      const controller = new AbortController();
      const timeoutId = window.setTimeout(() => controller.abort(), 6000);
      // eslint-disable-next-line no-await-in-loop
      const res = await fetch(url, { method: "GET", signal: controller.signal });
      window.clearTimeout(timeoutId);
      if (res.ok) return true;
    } catch {
      /* tenta o próximo candidato */
    }
  }
  return false;
}

export function ApiStatusBadge() {
  const [state, setState] = useState<ApiState>("checking");
  const lastState = useRef<ApiState>("checking");
  const cancelledRef = useRef(false);

  const ping = useCallback(async (manual: boolean) => {
    setState("checking");
    const ok = await checkHealth();
    if (cancelledRef.current) return;
    const next: ApiState = ok ? "online" : "offline";
    setState(next);
    if (manual || lastState.current !== next) {
      pushToast(ok ? "API Online" : "API Offline", ok ? "success" : "danger");
    }
    lastState.current = next;
  }, []);

  useEffect(() => {
    cancelledRef.current = false;
    ping(false);
    const interval = window.setInterval(() => ping(false), 45000);
    return () => {
      cancelledRef.current = true;
      window.clearInterval(interval);
    };
  }, [ping]);

  return (
    <button
      id="apiStatusBtn"
      className={`btn btn-outline-secondary api-status-btn api-${state}`}
      data-bs-toggle="tooltip"
      data-bs-title="Status da API"
      aria-label="Status da API"
      onClick={() => ping(true)}
    >
      <svg
        id="apiStatusIcon"
        xmlns="http://www.w3.org/2000/svg"
        className={`h-5 w-5${state === "checking" ? " spin-rotating" : ""}`}
        fill="none"
        viewBox="0 0 24 24"
        stroke="currentColor"
        strokeWidth={1.8}
        aria-hidden="true"
      >
        <StatusIcon state={state} />
      </svg>
      <span id="apiStatusText">{TEXT[state]}</span>
    </button>
  );
}
