import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "./index.css";
import "./toast/legacyBridge"; // efeito colateral: instala window.showToast
import "./history/historyStore"; // efeito colateral: instala window.addToHistory
import { App } from "./App";

const root = document.getElementById("root");
if (root) {
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}
