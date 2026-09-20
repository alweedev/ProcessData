import type { ReactNode } from "react";
import type { View } from "../routing/useHashRoute";
import { IconClock, IconHome, IconSitemap, IconUpload, IconUserMinus } from "../ui/icons";

export interface NavItem {
  id: View;
  label: string;
  icon: ReactNode;
}

export interface NavGroup {
  id: string;
  label: string;
  items: NavItem[];
}

/** Fonte única da navegação: a Sidebar renderiza os grupos e a TopBar deriva
 *  o breadcrumb daqui, então rótulo/ícone/agrupamento não divergem. */
export const NAV_GROUPS: NavGroup[] = [
  {
    id: "geral",
    label: "Geral",
    items: [{ id: "home", label: "Início", icon: <IconHome className="h-4 w-4" /> }],
  },
  {
    id: "operacoes",
    label: "Operações",
    items: [
      { id: "cadastro", label: "Cadastro", icon: <IconUpload className="h-4 w-4" /> },
      { id: "inativacao", label: "Inativação", icon: <IconUserMinus className="h-4 w-4" /> },
      { id: "estruturas", label: "Estruturas de aprovação", icon: <IconSitemap className="h-4 w-4" /> },
    ],
  },
  {
    id: "consulta",
    label: "Consulta",
    items: [{ id: "historico", label: "Histórico", icon: <IconClock className="h-4 w-4" /> }],
  },
];

export function findNav(view: View): { group: NavGroup; item: NavItem } {
  for (const group of NAV_GROUPS) {
    const item = group.items.find((candidate) => candidate.id === view);
    if (item) return { group, item };
  }
  // View desconhecida cai no Início (mesmo fallback do parseHash).
  return { group: NAV_GROUPS[0], item: NAV_GROUPS[0].items[0] };
}
