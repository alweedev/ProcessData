import { useEffect, useState } from "react";
import { IconMoon, IconSun } from "../ui/icons";
import { cn } from "../ui/cn";

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
    const commit = () => {
      setMode(next);
      try {
        localStorage.setItem(THEME_KEY, next);
      } catch {
        /* noop */
      }
    };

    // Crossfade da página inteira via View Transitions API nativa (sem
    // dependência) quando disponível; cai para a transição CSS de cor
    // (index.css) em navegadores sem suporte ou com prefers-reduced-motion.
    const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
    if (!reduced && typeof document.startViewTransition === "function") {
      document.startViewTransition(commit);
    } else {
      commit();
    }
  }

  const isDark = mode === "dark";
  return (
    <button
      id="themeToggle"
      type="button"
      role="switch"
      aria-checked={isDark}
      aria-label={isDark ? "Alternar para tema claro" : "Alternar para tema escuro"}
      title="Alternar tema"
      onClick={toggle}
      className={cn(
        "relative inline-flex h-8 w-14 shrink-0 items-center rounded-pill border transition-colors outline-none focus-visible:ring-2 focus-visible:ring-accent/50",
        isDark ? "border-brand/40 bg-brand-soft" : "border-border-strong bg-surface-2",
      )}
    >
      <span
        className={cn(
          "absolute left-1 flex h-6 w-6 items-center justify-center rounded-pill bg-surface text-text shadow-card transition-transform duration-300",
          isDark && "translate-x-6",
        )}
      >
        <IconSun
          className={cn(
            "absolute h-3.5 w-3.5 transition-all duration-300",
            isDark ? "-rotate-90 scale-50 opacity-0" : "rotate-0 scale-100 opacity-100",
          )}
        />
        <IconMoon
          className={cn(
            "absolute h-3.5 w-3.5 transition-all duration-300",
            isDark ? "rotate-0 scale-100 opacity-100" : "rotate-90 scale-50 opacity-0",
          )}
        />
      </span>
    </button>
  );
}
