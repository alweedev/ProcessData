import type { ReactNode } from "react";
import { Credits } from "./chrome/Credits";
import { SkipLink } from "./chrome/SkipLink";
import { Sidebar } from "./chrome/Sidebar";
import { TopBar } from "./chrome/TopBar";
import { ApiHealthBanner } from "./health/ApiHealthBanner";
import { useHashRoute, type View } from "./routing/useHashRoute";
import { CadastroTab } from "./tabs/CadastroTab";
import { EstruturasTab } from "./tabs/EstruturasTab";
import { HistoricoTab } from "./tabs/HistoricoTab";
import { InativacaoTab } from "./tabs/InativacaoTab";
import { ToastViewport } from "./toast/ToastViewport";
import { HomeView } from "./views/Home";

type TabId = Exclude<View, "home">;

const PANELS: { id: TabId; content: ReactNode }[] = [
  { id: "cadastro", content: <CadastroTab /> },
  { id: "inativacao", content: <InativacaoTab /> },
  { id: "estruturas", content: <EstruturasTab /> },
  { id: "historico", content: <HistoricoTab /> },
];

/**
 * Casca da aplicação em três áreas, cada elemento no seu lugar:
 *  - Sidebar (esquerda, altura total): marca, navegação agrupada, versão/créditos.
 *  - TopBar (topo do conteúdo): onde estou + status da API + tema.
 *  - Conteúdo: fluido e alinhado à esquerda (mesma margem do breadcrumb),
 *    com teto só para telas ultrawide.
 * No mobile as áreas empilham: TopBar (com a marca), navegação em faixa
 * horizontal e conteúdo — a ordem do DOM já é essa; o grid só reposiciona no md+.
 */
export function App() {
  const view = useHashRoute();

  return (
    <>
      <SkipLink />
      <div className="grid min-h-screen grid-cols-1 bg-bg text-text md:grid-cols-[15rem_minmax(0,1fr)] md:grid-rows-[auto_1fr]">
        <TopBar view={view} />
        <Sidebar view={view} />

        <div className="flex min-w-0 flex-col md:col-start-2 md:row-start-2">
          <ApiHealthBanner />

          <main id="conteudo" tabIndex={-1} className="w-full max-w-[1600px] flex-1 px-4 py-6 outline-none md:px-8 md:py-8">
            {view === "home" && <HomeView />}
            {PANELS.map(({ id, content }) => (
              <div
                key={id}
                id={id}
                role="tabpanel"
                aria-labelledby={`${id}-tab`}
                inert={view !== id}
                className={view !== id ? "pd-tabpanel-hidden" : undefined}
              >
                {content}
              </div>
            ))}
          </main>

          {/* Mobile: a Sidebar (com os créditos) some do fluxo, então eles descem aqui. */}
          <footer className="border-t border-border px-4 py-4 md:hidden">
            <Credits className="justify-center" />
          </footer>
        </div>
      </div>

      <div id="toastContainer" aria-live="polite" aria-atomic="false" className="fixed bottom-0 right-0 z-[1100] p-3">
        <ToastViewport />
      </div>
    </>
  );
}
