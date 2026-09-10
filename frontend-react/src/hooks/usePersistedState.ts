import { useState } from "react";

/**
 * Estado sincronizado com uma chave do localStorage (merge raso — cada
 * campo do objeto persiste independentemente, como os *_prefs legados).
 */
export function usePersistedState<T extends Record<string, unknown>>(key: string, initial: T) {
  const [state, setState] = useState<T>(() => {
    try {
      const stored = JSON.parse(localStorage.getItem(key) || "{}");
      return { ...initial, ...stored };
    } catch {
      return initial;
    }
  });

  function update(patch: Partial<T>) {
    setState((prev) => {
      const next = { ...prev, ...patch };
      try {
        localStorage.setItem(key, JSON.stringify(next));
      } catch {
        /* noop */
      }
      return next;
    });
  }

  return [state, update] as const;
}
