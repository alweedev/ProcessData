import { useSyncExternalStore } from "react";
import { flushSync } from "react-dom";

export type View = "home" | "cadastro" | "inativacao" | "estruturas" | "historico";

const VIEWS: readonly View[] = ["home", "cadastro", "inativacao", "estruturas", "historico"];

function parseHash(): View {
  const raw = window.location.hash.replace(/^#\/?/, "").trim().toLowerCase();
  return (VIEWS as readonly string[]).includes(raw) ? (raw as View) : "home";
}

type Listener = () => void;
const listeners = new Set<Listener>();
let view: View = parseHash();

function setView(next: View) {
  if (next === view) return;
  view = next;
  for (const listener of listeners) listener();
}

window.addEventListener("hashchange", () => setView(parseHash()));

/** Navega trocando o hash (`#/cadastro`). O router é só o hash — nunca toca
 *  no path, então o Flask nunca vê essas URLs e os testes e2e (que fazem
 *  `goto("/")` + clique no trigger) não são afetados.
 *
 *  Quando disponível (e sem prefers-reduced-motion), envolve a troca numa
 *  View Transition nativa — mesmo padrão do ThemeToggle — pra uma transição
 *  suave entre telas. `flushSync` garante que o React já comitou a nova
 *  view antes do browser tirar o "screenshot" de depois; sem isso a
 *  transição captura o estado antigo nos dois lados. A classe `vt-nav` em
 *  `<html>` deixa o CSS (index.css) diferenciar essa transição da do
 *  ThemeToggle (crossfade simples), removida quando a transição termina. */
export function navigate(next: View): void {
  const target = next === "home" ? "#/" : `#/${next}`;
  const commit = () => {
    if (window.location.hash !== target) window.location.hash = target;
    setView(next);
  };

  const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
  if (!reduced && typeof document.startViewTransition === "function") {
    document.documentElement.classList.add("vt-nav");
    const transition = document.startViewTransition(() => flushSync(commit));
    transition.finished.finally(() => document.documentElement.classList.remove("vt-nav"));
  } else {
    commit();
  }
}

function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function getSnapshot(): View {
  return view;
}

export function useHashRoute(): View {
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}
