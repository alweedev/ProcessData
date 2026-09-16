import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "./index.css";
import "./history/historyStore"; // efeito colateral: migra a chave legada do localStorage o quanto antes
import { App } from "./App";

const root = document.getElementById("root");
if (root) {
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}
