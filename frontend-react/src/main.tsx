import type { ComponentType } from "react";
import { createRoot } from "react-dom/client";
import "./index.css";
import "./toast/legacyBridge"; // efeito colateral: instala window.showToast
import "./history/historyStore"; // efeito colateral: instala window.addToHistory
import { AppChrome } from "./chrome/AppChrome";
import { CadastroTab } from "./tabs/CadastroTab";
import { HistoricoTab } from "./tabs/HistoricoTab";
import { InativacaoTab } from "./tabs/InativacaoTab";
import { ToastViewport } from "./toast/ToastViewport";

/**
 * Cada aba/pedaço de chrome migrado soma uma entrada aqui. O elemento do
 * seletor continua existindo e sendo controlado pelo markup legado (ex.: o
 * <div class="tab-pane"> que o Bootstrap tab-JS mostra/esconde) — só os
 * FILHOS são substituídos pelo React.
 */
interface Mount {
  selector: string;
  Component: ComponentType;
}

const MOUNTS: Mount[] = [
  { selector: "#appChromeControls", Component: AppChrome },
  { selector: "#toastContainer", Component: ToastViewport },
  { selector: "#historico", Component: HistoricoTab },
  { selector: "#inativacao", Component: InativacaoTab },
  { selector: "#cadastro", Component: CadastroTab },
];

function mount() {
  for (const { selector, Component } of MOUNTS) {
    const el = document.querySelector(selector);
    if (!el) continue;
    el.innerHTML = "";
    createRoot(el).render(<Component />);
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", mount);
} else {
  mount();
}
