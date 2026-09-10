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
  const isDark = mode === "dark";
  // Mantém .dark + data-bs-theme em <body> (não só <html>): o CSS legado
  // (custom.css/animations.css) ainda usado pelas abas não migradas depende
  // de `body.dark`/`body[data-bs-theme]` até a limpeza final (Fase 6).
  document.body.classList.toggle("dark", isDark);
  document.body.setAttribute("data-bs-theme", isDark ? "dark" : "light");
  window.__applyRasterInvertToUploadZones?.();
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
      className="btn btn-dark"
      data-bs-toggle="tooltip"
      data-bs-title="Alternar tema"
      aria-label={isDark ? "Alternar para tema claro" : "Alternar para tema escuro"}
      onClick={toggle}
    >
      <span id="themeIcon" className="theme-icon-wrapper" aria-hidden="true">
        <svg
          id="iconSun"
          xmlns="http://www.w3.org/2000/svg"
          className={`h-5 w-5 theme-sun${isDark ? " d-none" : ""}`}
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth={1.6}
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M12 4V2m0 20v-2m8-8h2M2 12h2m13.657-6.343l1.414-1.414M4.929 19.071l1.414-1.414m0-11.314L4.93 4.93m13.657 13.657l1.414 1.414"
          />
          <circle cx="12" cy="12" r="4" />
        </svg>
        <svg
          id="iconMoon"
          xmlns="http://www.w3.org/2000/svg"
          className={`h-5 w-5 theme-moon${isDark ? "" : " d-none"}`}
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth={1.6}
        >
          <path strokeLinecap="round" strokeLinejoin="round" d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z" />
        </svg>
      </span>
    </button>
  );
}
