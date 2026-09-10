export {};

declare global {
  interface Window {
    /** Instalado por src/toast/legacyBridge.ts. */
    showToast?: (message: string, type?: string) => void;
    /** Instalado por src/history/historyStore.ts. */
    addToHistory?: (action: string) => void;
  }
}
