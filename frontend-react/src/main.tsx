import type { ComponentType } from "react";
import { createRoot } from "react-dom/client";
import "./index.css";

/**
 * Cada aba migrada soma uma entrada aqui. O elemento do seletor é o próprio
 * <div class="tab-pane" id="..."> que o Bootstrap tab-JS já controla (mostra/
 * esconde via classe) — só os FILHOS são substituídos pelo React, o nó
 * externo continua sendo gerenciado pelo markup legado enquanto ele existir.
 */
interface MigratedTab {
  selector: string;
  Component: ComponentType;
}

const MIGRATED_TABS: MigratedTab[] = [
  // Fase 0 (AppShell) e as fases seguintes somam entradas aqui.
];

function mountTabs() {
  for (const { selector, Component } of MIGRATED_TABS) {
    const el = document.querySelector(selector);
    if (!el) continue;
    el.innerHTML = "";
    createRoot(el).render(<Component />);
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", mountTabs);
} else {
  mountTabs();
}
