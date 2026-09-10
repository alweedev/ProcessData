import { useState } from "react";
import { AppChrome } from "./chrome/AppChrome";
import { CadastroTab } from "./tabs/CadastroTab";
import { EstruturasTab } from "./tabs/EstruturasTab";
import { HistoricoTab } from "./tabs/HistoricoTab";
import { InativacaoTab } from "./tabs/InativacaoTab";
import { ToastViewport } from "./toast/ToastViewport";

type TabId = "cadastro" | "inativacao" | "estruturas" | "historico";

const TABS: { id: TabId; label: string }[] = [
  { id: "cadastro", label: "Cadastro" },
  { id: "inativacao", label: "Inativação" },
  { id: "estruturas", label: "Estruturas de aprovação" },
  { id: "historico", label: "Histórico" },
];

export function App() {
  const [active, setActive] = useState<TabId>("cadastro");

  return (
    <>
      <div className="mx-auto max-w-5xl px-4 py-6">
        <header className="mb-6 flex items-center justify-between gap-4">
          <button
            id="brand"
            type="button"
            onClick={() => setActive("cadastro")}
            className="flex items-center gap-2"
          >
            <span
              aria-hidden="true"
              className="flex h-9 w-9 items-center justify-center rounded-lg bg-gradient-to-br from-brand-from to-brand-to text-white"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="currentColor">
                <path d="M12 2a2 2 0 0 1 2 2v1h3a3 3 0 0 1 3 3v9a3 3 0 0 1-3 3H7a3 3 0 0 1-3-3V8a3 3 0 0 1 3-3h3V4a2 2 0 0 1 2-2Zm-3 9a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3Zm6 0a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3Z" />
              </svg>
            </span>
            <span className="text-xl font-bold text-accent dark:text-accent-dark">ProcessData</span>
          </button>
          <div className="flex items-center gap-2">
            <AppChrome />
          </div>
        </header>

        <nav id="myTab" role="tablist" className="mb-4 flex flex-wrap gap-1 border-b border-black/10 dark:border-white/10">
          {TABS.map((tab) => (
            <button
              key={tab.id}
              id={`${tab.id}-tab`}
              type="button"
              role="tab"
              aria-controls={tab.id}
              aria-selected={active === tab.id}
              onClick={() => setActive(tab.id)}
              className={`-mb-px rounded-t-lg border border-b-0 px-4 py-2 text-sm font-medium ${
                active === tab.id
                  ? "border-black/10 bg-surface text-slate-900 dark:border-white/10 dark:bg-surface-dark dark:text-slate-100"
                  : "border-transparent text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200"
              }`}
            >
              {tab.label}
            </button>
          ))}
        </nav>

        <div id="cadastro" role="tabpanel" aria-labelledby="cadastro-tab" hidden={active !== "cadastro"}>
          <CadastroTab />
        </div>
        <div id="inativacao" role="tabpanel" aria-labelledby="inativacao-tab" hidden={active !== "inativacao"}>
          <InativacaoTab />
        </div>
        <div id="estruturas" role="tabpanel" aria-labelledby="estruturas-tab" hidden={active !== "estruturas"}>
          <EstruturasTab />
        </div>
        <div id="historico" role="tabpanel" aria-labelledby="historico-tab" hidden={active !== "historico"}>
          <HistoricoTab />
        </div>

        <footer className="mt-10 flex flex-wrap items-center justify-center gap-3 text-xs text-slate-500 dark:text-slate-400">
          <span>© 2026 ProcessData</span>
          <span className="rounded-full bg-black/5 px-2 py-0.5 dark:bg-white/10">v1.0.0</span>
          <span>
            Desenvolvido por{" "}
            <a
              href="https://www.linkedin.com/in/alejandro-gabriel/"
              target="_blank"
              rel="noopener noreferrer"
              className="text-accent hover:underline dark:text-accent-dark"
            >
              Alejandro Gabriel
            </a>
          </span>
        </footer>
      </div>

      <div id="toastContainer" aria-live="polite" aria-atomic="false" className="fixed bottom-0 right-0 z-[1100] p-3">
        <ToastViewport />
      </div>
    </>
  );
}
