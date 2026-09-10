const HISTORY_KEY = "history_v2";
const LEGACY_HISTORY_KEY = "history";
const API_URL = "/api/history";

export interface HistoryEntry {
  ts: number;
  text: string;
  source: "local" | "server";
  details?: Record<string, unknown>;
}

export interface HistoryState {
  merged: HistoryEntry[];
  localCount: number;
  serverCount: number;
}

type Listener = (state: HistoryState) => void;

let serverEntries: HistoryEntry[] = [];
const listeners = new Set<Listener>();

function migrateLegacyKey() {
  try {
    if (localStorage.getItem(HISTORY_KEY)) return;
    const legacy = JSON.parse(localStorage.getItem(LEGACY_HISTORY_KEY) || "[]");
    if (!Array.isArray(legacy) || !legacy.length) return;
    const now = Date.now();
    const migrated: HistoryEntry[] = legacy.map((text: unknown, index: number) => ({
      ts: now - (legacy.length - index) * 1000,
      text: String(text),
      source: "local",
    }));
    localStorage.setItem(HISTORY_KEY, JSON.stringify(migrated.slice(-200)));
    localStorage.removeItem(LEGACY_HISTORY_KEY);
  } catch {
    /* noop */
  }
}

function getLocalEntries(): HistoryEntry[] {
  try {
    const arr = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function setLocalEntries(arr: HistoryEntry[]) {
  try {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(arr.slice(-500)));
  } catch {
    /* noop */
  }
}

function computeState(): HistoryState {
  const local = getLocalEntries();
  const merged = [...local, ...serverEntries].sort((a, b) => b.ts - a.ts).slice(0, 500);
  return { merged, localCount: local.length, serverCount: serverEntries.length };
}

function emit() {
  const state = computeState();
  for (const listener of listeners) listener(state);
}

export function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  listener(computeState());
  return () => {
    listeners.delete(listener);
  };
}

export function getSnapshot(): HistoryState {
  return computeState();
}

export async function refreshFromServer(): Promise<void> {
  try {
    const res = await fetch(`${API_URL}?limit=300`);
    if (!res.ok) return;
    const data = await res.json();
    if (!data || !Array.isArray(data.items)) return;
    serverEntries = data.items.map(
      (item: { timestamp?: string; event_type?: string; status?: string; details?: Record<string, unknown> }) => ({
        ts: Date.parse(item.timestamp || "") || Date.now(),
        text: `${item.event_type || "evento"}: ${item.status || "status"}`,
        source: "server" as const,
        details: item.details || {},
      }),
    );
    emit();
  } catch {
    /* noop */
  }
}

export function addLocalEntry(text: string): void {
  const entries = getLocalEntries();
  entries.push({ ts: Date.now(), text: String(text), source: "local" });
  setLocalEntries(entries);
  emit();
}

export async function clearAll(): Promise<void> {
  setLocalEntries([]);
  try {
    await fetch(API_URL, { method: "DELETE" });
  } catch {
    /* noop */
  }
  serverEntries = [];
  emit();
}

migrateLegacyKey();

/**
 * Abas ainda não migradas (Cadastro, Análise, Estruturas) continuam chamando
 * `window.addToHistory(texto)` normalmente — importar este módulo já instala
 * a ponte, mesmo padrão de src/toast/legacyBridge.ts.
 */
window.addToHistory = addLocalEntry;
