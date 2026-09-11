import { useSyncExternalStore } from "react";

export type RunOperation = "cadastro" | "inativacao" | "estruturas";

export interface RunEntry {
  id: string;
  operation: RunOperation;
  ts: number;
  inputSummary: string[];
  outputFilename: string;
  /** URL do blob retida na sessão para "baixar novamente". */
  blobUrl?: string;
}

type Listener = (runs: RunEntry[]) => void;

let runs: RunEntry[] = [];
const listeners = new Set<Listener>();

function emit() {
  for (const listener of listeners) listener(runs);
}

/** Registra uma execução bem-sucedida (só em memória, sem backend). */
export function addRun(entry: {
  operation: RunOperation;
  inputSummary: string[];
  outputFilename: string;
  blobUrl?: string;
  ts?: number;
}): RunEntry {
  const run: RunEntry = {
    id: `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    ts: entry.ts ?? Date.now(),
    operation: entry.operation,
    inputSummary: entry.inputSummary,
    outputFilename: entry.outputFilename,
    blobUrl: entry.blobUrl,
  };
  runs = [run, ...runs].slice(0, 50);
  emit();
  return run;
}

function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

function getSnapshot(): RunEntry[] {
  return runs;
}

export function useRuns(): RunEntry[] {
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}
