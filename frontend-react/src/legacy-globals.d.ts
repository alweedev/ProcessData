export {};

declare global {
  interface Window {
    /** Instalado por src/toast/legacyBridge.ts, consumido pelo JS legado ainda não migrado. */
    showToast?: (message: string, type?: string) => void;
    /** Definido em frontend/static/js/app.v2.js; chamado após alternar o tema. */
    __applyRasterInvertToUploadZones?: () => void;
    /** Definido em frontend/static/js/history/index.js. */
    addToHistory?: (action: string) => void;
    /** Prefixo opcional para chamadas /api/* (definido pelo app legado, se houver). */
    API_BASE?: string;
  }
}
