import { useEffect, useState } from "react";

const MOTION_KEY = "motion_pref";

function computeInitialReduced(): boolean {
  try {
    const stored = localStorage.getItem(MOTION_KEY);
    if (stored) return stored === "reduce";
  } catch {
    /* noop */
  }
  return window.matchMedia?.("(prefers-reduced-motion: reduce)").matches ?? false;
}

export function MotionToggle() {
  const [reduced, setReduced] = useState<boolean>(computeInitialReduced);

  useEffect(() => {
    document.body.classList.toggle("reduce-motion", reduced);
  }, [reduced]);

  function toggle() {
    const next = !reduced;
    setReduced(next);
    try {
      localStorage.setItem(MOTION_KEY, next ? "reduce" : "full");
    } catch {
      /* noop */
    }
  }

  return (
    <button
      id="motionToggle"
      className="btn btn-outline-secondary"
      data-bs-toggle="tooltip"
      data-bs-title="Reduzir movimentos"
      aria-label="Reduzir movimentos"
      style={{ display: "inline-flex", alignItems: "center", gap: "0.5rem" }}
      onClick={toggle}
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        className="h-5 w-5"
        fill="none"
        viewBox="0 0 24 24"
        stroke="currentColor"
        strokeWidth={1.8}
        aria-hidden="true"
        style={{ flexShrink: 0 }}
      >
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          d="M12 3C7.03 3 3 7.03 3 12c0 1.76.57 3.39 1.53 4.72L16.72 4.53A8.963 8.963 0 0012 3z"
        />
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          d="M21 12a8.963 8.963 0 00-1.53-4.72L7.28 19.47A8.963 8.963 0 0012 21c4.97 0 9-4.03 9-9z"
        />
      </svg>
      <span>{reduced ? "Animações: Reduzidas" : "Animações: Ativas"}</span>
    </button>
  );
}
