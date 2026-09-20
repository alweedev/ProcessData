import type { View } from "../routing/useHashRoute";
import { ApiStatus } from "./ApiStatus";
import { Brand } from "./Brand";
import { findNav } from "./navConfig";
import { ThemeToggle } from "./ThemeToggle";

export function TopBar({ view }: { view: View }) {
  const { group, item } = findNav(view);

  return (
    <header className="sticky top-0 z-30 flex h-14 items-center justify-between gap-4 border-b border-border bg-surface/85 px-4 backdrop-blur md:col-start-2 md:row-start-1 md:px-8">
      {/* Mobile: marca aqui (no desktop ela mora na Sidebar). */}
      <Brand className="md:hidden" />

      {/* Desktop: onde o usuário está — a marca e o título da página já têm
          seus próprios lugares, então aqui só o contexto. */}
      <nav aria-label="Você está em" className="hidden min-w-0 text-sm md:block">
        <ol className="flex items-center gap-2 text-text-muted">
          <li>{view === "home" ? "ProcessData" : group.label}</li>
          <li aria-hidden="true" className="text-text-subtle">
            /
          </li>
          <li aria-current="page" className="truncate font-medium text-text">
            {view === "home" ? "Visão geral" : item.label}
          </li>
        </ol>
      </nav>

      <div className="flex shrink-0 items-center gap-3">
        <ApiStatus />
        <ThemeToggle />
      </div>
    </header>
  );
}
