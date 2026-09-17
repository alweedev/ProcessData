export const THEME_KEY = "pd_theme";

export type Mode = "light" | "dark";

/** Tema padrão quando não há preferência salva: dark, a menos que o SO peça
 *  explicitamente light (prefers-color-scheme: light). Mesma lógica do
 *  script anti-FOUC em frontend/index.html — os dois precisam ficar em
 *  sincronia pra não haver flash de tema errado no 1º paint. */
export function computeInitialTheme(): Mode {
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === "dark" || stored === "light") return stored;
  } catch {
    /* localStorage indisponível (modo privado etc.) */
  }
  const prefersLight = window.matchMedia?.("(prefers-color-scheme: light)").matches;
  return prefersLight ? "light" : "dark";
}
