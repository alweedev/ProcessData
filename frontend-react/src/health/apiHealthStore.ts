import { useSyncExternalStore } from "react";
import { pushToast } from "../toast/toastStore";

export type ApiState = "online" | "offline" | "checking";

type Listener = (state: ApiState) => void;

let state: ApiState = "checking";
let lastResolved: ApiState = "checking";
const listeners = new Set<Listener>();
let started = false;
let intervalId: number | undefined;

function setState(next: ApiState) {
  if (next === state) return;
  state = next;
  for (const listener of listeners) listener(state);
}

async function probe(): Promise<boolean> {
  for (const url of ["/api/health", "/"]) {
    try {
      const controller = new AbortController();
      const timeoutId = window.setTimeout(() => controller.abort(), 6000);
      const res = await fetch(url, { method: "GET", signal: controller.signal });
      window.clearTimeout(timeoutId);
      if (res.ok) return true;
    } catch {
      /* tenta o próximo candidato */
    }
  }
  return false;
}

/** Faz um ping imediato. `manual=true` sempre toasta o resultado; caso
 *  contrário só toasta em transições reais de estado — a primeira resolução
 *  da sessão não toasta quando cai em "online" (caminho feliz esperado no
 *  carregamento da página), só quando já cai "offline" ou quando o estado
 *  muda depois de já ter resolvido uma vez. */
export async function checkApiHealth(manual = false): Promise<void> {
  setState("checking");
  const ok = await probe();
  const next: ApiState = ok ? "online" : "offline";
  setState(next);
  const isFirstResolution = lastResolved === "checking";
  const changed = lastResolved !== next;
  if (manual || (changed && !(isFirstResolution && next === "online"))) {
    pushToast(ok ? "API Online" : "API Offline", ok ? "success" : "danger");
  }
  lastResolved = next;
}

function ensureStarted() {
  if (started) return;
  started = true;
  void checkApiHealth(false);
  intervalId = window.setInterval(() => void checkApiHealth(false), 45000);
}

function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  ensureStarted();
  return () => {
    listeners.delete(listener);
    if (listeners.size === 0 && intervalId !== undefined) {
      window.clearInterval(intervalId);
      intervalId = undefined;
      started = false;
    }
  };
}

function getSnapshot(): ApiState {
  return state;
}

export function useApiHealth(): ApiState {
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}
