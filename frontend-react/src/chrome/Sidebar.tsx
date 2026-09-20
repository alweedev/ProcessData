import { navigate, type View } from "../routing/useHashRoute";
import { Brand } from "./Brand";
import { Credits } from "./Credits";
import { NAV_GROUPS, type NavItem } from "./navConfig";

interface SidebarProps {
  view: View;
}

function itemClass(current: boolean): string {
  return [
    "relative flex shrink-0 items-center gap-2 rounded-control px-3 py-2 text-left text-sm font-medium outline-none transition-colors focus-visible:ring-2 focus-visible:ring-accent/50",
    // Barra de destaque à esquerda só no layout de coluna (md+).
    "md:before:absolute md:before:inset-y-1.5 md:before:left-0 md:before:w-0.5 md:before:rounded-full md:before:transition-colors",
    current
      ? "bg-accent/10 text-accent-text md:before:bg-accent"
      : "text-text-muted hover:bg-surface-2 hover:text-text md:before:bg-transparent",
  ].join(" ");
}

function NavButton({ item, current, tab }: { item: NavItem; current: boolean; tab: boolean }) {
  const common = {
    type: "button" as const,
    onClick: () => navigate(item.id),
    className: itemClass(current),
  };
  // Operações e Histórico são abas (controlam os tabpanels em App.tsx); o
  // Início é só um link de navegação para a view de visão geral.
  if (!tab) {
    return (
      <button {...common} aria-current={current ? "page" : undefined}>
        {item.icon}
        <span>{item.label}</span>
      </button>
    );
  }
  return (
    <button
      {...common}
      id={`${item.id}-tab`}
      role="tab"
      aria-controls={item.id}
      aria-selected={current}
      tabIndex={current ? 0 : -1}
    >
      {item.icon}
      <span>{item.label}</span>
    </button>
  );
}

export function Sidebar({ view }: SidebarProps) {
  return (
    <aside className="border-b border-border bg-surface md:sticky md:top-0 md:col-start-1 md:row-span-2 md:row-start-1 md:flex md:h-screen md:flex-col md:self-start md:border-b-0 md:border-r">
      {/* Marca no desktop; no mobile ela fica na TopBar. Altura = TopBar, para
          as duas bordas inferiores formarem uma linha contínua. */}
      <div className="hidden h-14 shrink-0 items-center border-b border-border px-4 md:flex">
        <Brand id="brand" />
      </div>

      <nav
        aria-label="Navegação"
        className="flex gap-1 overflow-x-auto px-3 py-2 md:flex-1 md:flex-col md:gap-5 md:overflow-y-auto md:px-3 md:py-4"
      >
        {NAV_GROUPS.map((group) => {
          const isTabGroup = group.id !== "geral";
          const labelId = `nav-${group.id}-label`;
          return (
            <div key={group.id} className="flex gap-1 md:flex-col">
              <p
                id={labelId}
                className="hidden px-3 pb-1 text-[11px] font-semibold uppercase tracking-wider text-text-subtle md:block"
              >
                {group.label}
              </p>
              <div
                id={group.id === "operacoes" ? "myTab" : undefined}
                role={isTabGroup ? "tablist" : "group"}
                aria-labelledby={labelId}
                className="flex gap-1 md:flex-col"
              >
                {group.items.map((item) => (
                  <NavButton key={item.id} item={item} current={view === item.id} tab={isTabGroup} />
                ))}
              </div>
            </div>
          );
        })}
      </nav>

      <div className="hidden shrink-0 border-t border-border px-4 py-3 md:block">
        <Credits />
      </div>
    </aside>
  );
}
