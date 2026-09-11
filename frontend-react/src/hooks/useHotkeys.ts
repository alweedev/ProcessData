import { useEffect } from "react";
import type { View } from "../routing/useHashRoute";

interface HotkeyHandlers {
  onNavigate: (view: View) => void;
  onToggleHelp: () => void;
  onRunPrimary: () => void;
}

function isTypingTarget(el: EventTarget | null): boolean {
  if (!(el instanceof HTMLElement)) return false;
  const tag = el.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || el.isContentEditable;
}

const NAV_KEYS: Record<string, View> = {
  h: "home",
  "1": "cadastro",
  "2": "inativacao",
  "3": "estruturas",
  "4": "historico",
};

/** Atalhos de nível de casca, sem dependência. Navegação por tecla é ignorada
 *  enquanto o foco está num campo de texto; Ctrl/Cmd+Enter funciona sempre. */
export function useHotkeys({ onNavigate, onToggleHelp, onRunPrimary }: HotkeyHandlers): void {
  useEffect(() => {
    function onKeyDown(e: KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        onRunPrimary();
        return;
      }
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      if (isTypingTarget(e.target)) return;

      if (e.key === "?" || (e.shiftKey && e.key === "/")) {
        e.preventDefault();
        onToggleHelp();
        return;
      }
      const target = NAV_KEYS[e.key.toLowerCase()];
      if (target) {
        e.preventDefault();
        onNavigate(target);
      }
    }
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [onNavigate, onToggleHelp, onRunPrimary]);
}
