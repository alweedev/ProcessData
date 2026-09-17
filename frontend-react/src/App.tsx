import type { ReactNode } from "react";
import { AppChrome } from "./chrome/AppChrome";
import { ApiHealthBanner } from "./health/ApiHealthBanner";
import { navigate, useHashRoute, type View } from "./routing/useHashRoute";
import { CadastroTab } from "./tabs/CadastroTab";
import { EstruturasTab } from "./tabs/EstruturasTab";
import { HistoricoTab } from "./tabs/HistoricoTab";
import { InativacaoTab } from "./tabs/InativacaoTab";
import { ToastViewport } from "./toast/ToastViewport";
import { Badge } from "./ui/Badge";
import { IconClock, IconHome, IconSitemap, IconUpload, IconUserMinus } from "./ui/icons";
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
    "flex shrink-0 items-center gap-2 rounded-control px-3 py-2 text-left text-sm font-medium transition-colors",
    current ? "bg-accent/10 text-accent" : "text-text-muted hover:bg-surface-2 hover:text-text",
  ].join(" ");
}

export function App() {
  const view = useHashRoute();
  const active: TabId | null = view === "home" ? null : view;

  return (
    <>
      <div className="min-h-screen bg-bg text-text">
        <header className="sticky top-0 z-30 border-b border-border bg-surface/85 backdrop-blur">
          <div className="mx-auto flex max-w-[1360px] items-center justify-between gap-4 px-4 py-3">
            <button id="brand" type="button" onClick={() => navigate("home")} className="flex items-center gap-2">
              <span
                aria-hidden="true"
                className="flex h-9 w-9 items-center justify-center rounded-control bg-gradient-to-br from-brand-from to-brand-to text-white"
              >
                <span className="text-sm font-extrabold leading-none tracking-tighter">PD</span>
              </span>
              <span className="text-lg font-bold tracking-tight text-text">ProcessData</span>
            </button>
            <div className="flex items-center gap-2">
              <AppChrome />
            </div>
          </div>
        </header>

        <ApiHealthBanner />

        <div className="mx-auto flex max-w-[1360px] flex-col gap-6 px-4 py-6 md:flex-row">
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

        <footer className="border-t border-border">
          <div className="mx-auto flex max-w-[1360px] flex-wrap items-center justify-center gap-x-3 gap-y-1 px-4 pb-8 pt-4 text-xs text-text-subtle">
            <span>© 2026 ProcessData</span>
            <Badge tone="neutral">v{__APP_VERSION__}</Badge>
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
          </div>
        </footer>
      </div>

      <div id="toastContainer" aria-live="polite" aria-atomic="false" className="fixed bottom-0 right-0 z-[1100] p-3">
        <ToastViewport />
      </div>
    </>
  );
}
