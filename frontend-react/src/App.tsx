import { useCallback, useMemo, useRef, useState, type ReactNode } from "react";
import { ActionContext, type ActionRegistry } from "./app/actionContext";
import { AppChrome } from "./chrome/AppChrome";
import { ShortcutsHelp } from "./components/ShortcutsHelp";
import { useHotkeys } from "./hooks/useHotkeys";
import { navigate, useHashRoute, type View } from "./routing/useHashRoute";
import { CadastroTab } from "./tabs/CadastroTab";
import { EstruturasTab } from "./tabs/EstruturasTab";
import { HistoricoTab } from "./tabs/HistoricoTab";
import { InativacaoTab } from "./tabs/InativacaoTab";
import { ToastViewport } from "./toast/ToastViewport";
import { IconClock, IconHome, IconKeyboard, IconSitemap, IconUpload, IconUserMinus } from "./ui/icons";
import { HomeView } from "./views/Home";

type TabId = Exclude<View, "home">;

const TABS: { id: TabId; label: string; icon: ReactNode }[] = [
  { id: "cadastro", label: "Cadastro", icon: <IconUpload className="h-4 w-4" /> },
  { id: "inativacao", label: "Inativação", icon: <IconUserMinus className="h-4 w-4" /> },
  { id: "estruturas", label: "Estruturas de aprovação", icon: <IconSitemap className="h-4 w-4" /> },
  { id: "historico", label: "Histórico", icon: <IconClock className="h-4 w-4" /> },
];

function navItemClass(current: boolean): string {
  return [
    "flex shrink-0 items-center gap-2 rounded-lg px-3 py-2 text-left text-sm font-medium transition-colors",
    current ? "bg-accent/10 text-accent" : "text-text-muted hover:bg-surface-2 hover:text-text",
  ].join(" ");
}

export function App() {
  const view = useHashRoute();
  const active: TabId | null = view === "home" ? null : view;
  const [helpOpen, setHelpOpen] = useState(false);

  const actionRef = useRef<(() => void) | null>(null);
  const registry = useMemo<ActionRegistry>(
    () => ({
      register: (fn) => {
        actionRef.current = fn;
      },
      run: () => actionRef.current?.(),
    }),
    [],
  );

  const toggleHelp = useCallback(() => setHelpOpen((v) => !v), []);
  useHotkeys({ onNavigate: navigate, onToggleHelp: toggleHelp, onRunPrimary: registry.run });

  return (
    <ActionContext.Provider value={registry}>
      <div className="min-h-screen bg-bg text-text">
        <header className="sticky top-0 z-30 border-b border-border bg-surface/85 backdrop-blur">
          <div className="mx-auto flex max-w-6xl items-center justify-between gap-4 px-4 py-3">
            <button id="brand" type="button" onClick={() => navigate("home")} className="flex items-center gap-2">
              <span
                aria-hidden="true"
                className="flex h-9 w-9 items-center justify-center rounded-lg bg-gradient-to-br from-brand-from to-brand-to text-white"
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M12 2a2 2 0 0 1 2 2v1h3a3 3 0 0 1 3 3v9a3 3 0 0 1-3 3H7a3 3 0 0 1-3-3V8a3 3 0 0 1 3-3h3V4a2 2 0 0 1 2-2Zm-3 9a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3Zm6 0a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3Z" />
                </svg>
              </span>
              <span className="text-lg font-bold tracking-tight text-text">ProcessData</span>
            </button>
            <div className="flex items-center gap-2">
              <button
                id="shortcutsHelpBtn"
                type="button"
                aria-label="Atalhos de teclado"
                title="Atalhos de teclado (?)"
                onClick={toggleHelp}
                className="hidden rounded-lg border border-border-strong p-2 text-text-muted transition-colors hover:bg-surface-2 hover:text-text sm:inline-flex"
              >
                <IconKeyboard className="h-5 w-5" />
              </button>
              <AppChrome />
            </div>
          </div>
        </header>

        <div className="mx-auto flex max-w-6xl flex-col gap-6 px-4 py-6 md:flex-row">
          <aside className="md:w-56 md:shrink-0">
            <nav aria-label="Navegação" className="-mx-1 flex gap-1 overflow-x-auto px-1 md:mx-0 md:flex-col md:overflow-visible md:px-0">
              <button
                type="button"
                onClick={() => navigate("home")}
                aria-current={view === "home" ? "page" : undefined}
                className={navItemClass(view === "home")}
              >
                <IconHome className="h-4 w-4" />
                <span>Início</span>
              </button>
              <div id="myTab" role="tablist" aria-label="Operações" className="flex gap-1 md:flex-col">
                {TABS.map((tab) => (
                  <button
                    key={tab.id}
                    id={`${tab.id}-tab`}
                    type="button"
                    role="tab"
                    aria-controls={tab.id}
                    aria-selected={active === tab.id}
                    tabIndex={active === tab.id ? 0 : -1}
                    onClick={() => navigate(tab.id)}
                    className={navItemClass(active === tab.id)}
                  >
                    {tab.icon}
                    <span>{tab.label}</span>
                  </button>
                ))}
              </div>
            </nav>
          </aside>

          <main className="min-w-0 flex-1">
            {view === "home" && <HomeView />}
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
          </main>
        </div>

        <footer className="mx-auto flex max-w-6xl flex-wrap items-center justify-center gap-x-3 gap-y-1 px-4 pb-8 pt-2 text-xs text-text-subtle">
          <span>© 2026 ProcessData</span>
          <span className="rounded-full bg-surface-2 px-2 py-0.5">v1.0.0</span>
          <span>
            Desenvolvido por{" "}
            <a
              href="https://www.linkedin.com/in/alejandro-gabriel/"
              target="_blank"
              rel="noopener noreferrer"
              className="text-accent hover:underline"
            >
              Alejandro Gabriel
            </a>
          </span>
        </footer>
      </div>

      <div id="toastContainer" aria-live="polite" aria-atomic="false" className="fixed bottom-0 right-0 z-[1100] p-3">
        <ToastViewport />
      </div>

      <ShortcutsHelp open={helpOpen} onClose={() => setHelpOpen(false)} />
    </ActionContext.Provider>
  );
}
