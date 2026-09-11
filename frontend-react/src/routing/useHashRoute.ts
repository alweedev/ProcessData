import { useEffect, useState } from "react";

export type View = "home" | "cadastro" | "inativacao" | "estruturas" | "historico";

const VIEWS: readonly View[] = ["home", "cadastro", "inativacao", "estruturas", "historico"];

function parseHash(): View {
  const raw = window.location.hash.replace(/^#\/?/, "").trim().toLowerCase();
  return (VIEWS as readonly string[]).includes(raw) ? (raw as View) : "home";
}

/** Navega trocando o hash (`#/cadastro`). O router é só o hash — nunca toca no
 *  path, então o Flask nunca vê essas URLs e os testes e2e (que fazem
 *  `goto("/")` + clique no trigger) não são afetados. */
export function navigate(view: View): void {
  const target = view === "home" ? "#/" : `#/${view}`;
  if (window.location.hash !== target) window.location.hash = target;
}

export function useHashRoute(): View {
  const [view, setView] = useState<View>(parseHash);
  useEffect(() => {
    const onChange = () => setView(parseHash());
    window.addEventListener("hashchange", onChange);
    return () => window.removeEventListener("hashchange", onChange);
  }, []);
  return view;
}
