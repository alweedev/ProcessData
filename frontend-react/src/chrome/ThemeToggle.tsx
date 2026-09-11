import { useEffect, useState } from "react";

const THEME_KEY = "pd_theme";

type Mode = "light" | "dark";

function computeInitialTheme(): Mode {
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === "dark" || stored === "light") return stored;
  } catch {
    /* localStorage indisponível (modo privado etc.) */
  }
  const prefersDark = window.matchMedia?.("(prefers-color-scheme: dark)").matches;
  return prefersDark ? "dark" : "light";
}

function applyTheme(mode: Mode) {
  // .dark em <html> — mesmo alvo do script anti-FOUC no <head>.
  document.documentElement.classList.toggle("dark", mode === "dark");
}

export function ThemeToggle() {
  const [mode, setMode] = useState<Mode>(computeInitialTheme);

  useEffect(() => {
    applyTheme(mode);
  }, [mode]);

  function toggle() {
    const next: Mode = mode === "dark" ? "light" : "dark";
    setMode(next);
    try {
      localStorage.setItem(THEME_KEY, next);
    } catch {
      /* noop */
    }
  }

  const isDark = mode === "dark";
  return (
    <button
      id="themeToggle"
      type="button"
      aria-label={isDark ? "Alternar para tema claro" : "Alternar para tema escuro"}
      title="Alternar tema"
      onClick={toggle}
      className="rounded-lg border border-border-strong p-2 text-text-muted transition-colors hover:bg-surface-2 hover:text-text"
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        className="h-5 w-5"
        fill="none"
        viewBox="0 0 24 24"
        stroke="currentColor"
        strokeWidth={1.6}
        aria-hidden="true"
      >
        {isDark ? (
          <path strokeLinecap="round" strokeLinejoin="round" d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z" />
        ) : (
          <>
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M12 4V2m0 20v-2m8-8h2M2 12h2m13.657-6.343l1.414-1.414M4.929 19.071l1.414-1.414m0-11.314L4.93 4.93m13.657 13.657l1.414 1.414"
            />
            <circle cx="12" cy="12" r="4" />
          </>
        )}
      </svg>
    </button>
  );
}
