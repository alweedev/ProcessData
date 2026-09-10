export {};

declare global {
  interface Window {
    /** Instalado por src/toast/legacyBridge.ts. */
    showToast?: (message: string, type?: string) => void;
    /** Instalado por src/history/historyStore.ts. */
    addToHistory?: (action: string) => void;
    /** SweetAlert2, carregado via CDN (script global, sem tipos) até a Fase 6. */
    Swal?: {
      fire: (options: Record<string, unknown>) => Promise<unknown>;
    };
  }
}
