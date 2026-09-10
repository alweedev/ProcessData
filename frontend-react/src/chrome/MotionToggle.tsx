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
      type="button"
      aria-label="Reduzir movimentos"
      title="Reduzir movimentos"
      aria-pressed={reduced}
      onClick={toggle}
      className="inline-flex items-center gap-2 rounded-lg border border-black/15 px-3 py-2 text-sm text-slate-700 hover:bg-black/5 dark:border-white/15 dark:text-slate-200 dark:hover:bg-white/10"
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        className="h-5 w-5 shrink-0"
        fill="none"
        viewBox="0 0 24 24"
        stroke="currentColor"
        strokeWidth={1.8}
        aria-hidden="true"
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
