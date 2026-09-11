import { useEffect, useState } from "react";
import {
  getSnapshot as historySnapshot,
  refreshFromServer,
  subscribe as historySubscribe,
  type HistoryState,
} from "../../history/historyStore";
import { useRuns, type RunOperation } from "../../runs/runsStore";

export interface RecentItem {
  key: string;
  ts: number;
  title: string;
  subtitle?: string;
  outputFilename?: string;
  status: "success" | "error";
}

const OP_LABEL: Record<RunOperation, string> = {
  cadastro: "Cadastro",
  inativacao: "Inativação",
  estruturas: "Estruturas de aprovação",
};

/** event_type do audit log (backend/services/audit_service.py) que representam uma
 *  operação concluída (arquivo gerado) — o que faz sentido mostrar como "execução"
 *  no dashboard. Eventos de busca/validação (inativacao_busca, analysis_summary)
 *  ficam de fora daqui, mas continuam visíveis na íntegra na aba Histórico. */
const HISTORY_EVENT_LABEL: Record<string, string> = {
  cadastro: "Cadastro em massa",
  inativacao_geracao: "Inativação",
  aprovacao_remocao: "Estruturas de aprovação",
};

/** Funde as execuções ricas da sessão (runsStore) com o histórico
 *  local+servidor (historyStore), remove entradas de histórico que casam com
 *  uma execução da sessão (janela de 3s) e devolve as `limit` mais recentes. */
export function useRecentRuns(limit = 6): RecentItem[] {
  const runs = useRuns();
  const [history, setHistory] = useState<HistoryState>(historySnapshot);

  useEffect(() => {
    const unsubscribe = historySubscribe(setHistory);
    void refreshFromServer();
    return unsubscribe;
  }, []);

  const fromRuns: RecentItem[] = runs.map((run) => ({
    key: `run-${run.id}`,
    ts: run.ts,
    title: OP_LABEL[run.operation],
    subtitle: run.inputSummary.join(" · ") || undefined,
    outputFilename: run.outputFilename,
    status: "success",
  }));

  const runSeconds = new Set(fromRuns.map((item) => Math.round(item.ts / 3000)));
  const fromHistory: RecentItem[] = history.merged
    .filter((entry) => !runSeconds.has(Math.round(entry.ts / 3000)))
    .map((entry) => {
      const [eventType, status] = entry.text.split(":").map((s) => s.trim());
      const label = eventType ? HISTORY_EVENT_LABEL[eventType] : undefined;
      return { entry, label, status };
    })
    .filter((item): item is { entry: (typeof history.merged)[number]; label: string; status: string } =>
      Boolean(item.label),
    )
    .map(({ entry, label, status }) => ({
      key: `hist-${entry.source}-${entry.ts}`,
      ts: entry.ts,
      title: label,
      status: status === "success" ? "success" : "error",
    }));

  return [...fromRuns, ...fromHistory].sort((a, b) => b.ts - a.ts).slice(0, limit);
}
