import { useEffect, useMemo, useState } from "react";
import { formatTs } from "../../history/formatTs";
import { clearAll, getSnapshot, refreshFromServer, subscribe, type HistoryState } from "../../history/historyStore";
import { downloadFile } from "../../lib/downloadFile";
import { pushToast } from "../../toast/toastStore";

export function HistoricoTab() {
  const [state, setState] = useState<HistoryState>(getSnapshot);
  const [search, setSearch] = useState("");

  useEffect(() => subscribe(setState), []);
  useEffect(() => {
    refreshFromServer();
  }, []);

  const filtered = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return state.merged;
    return state.merged.filter((item) => item.text.toLowerCase().includes(term));
  }, [state.merged, search]);

  function exportCsv() {
    if (!filtered.length) {
      pushToast("Nada para exportar.", "danger");
      return;
    }
    const header = "data_hora,origem,acao";
    const rows = filtered.map((h) => {
      const action = h.text.replace(/"/g, '""');
      return `${formatTs(h.ts).replace(/,/g, " ")},${h.source},"${action}"`;
    });
    downloadFile([header, ...rows].join("\n"), `historico_${Date.now()}.csv`, "text/csv;charset=utf-8");
    pushToast("CSV exportado.", "success");
  }

  function exportJson() {
    if (!filtered.length) {
      pushToast("Nada para exportar.", "danger");
      return;
    }
    downloadFile(JSON.stringify(filtered, null, 2), `historico_${Date.now()}.json`, "application/json");
    pushToast("JSON exportado.", "success");
  }

  async function handleClear() {
    await clearAll();
    pushToast("Histórico limpo!", "success");
  }

  return (
    <div className="rounded-xl border border-black/10 bg-surface p-4 shadow-sm dark:border-white/10 dark:bg-surface-dark sm:p-6">
      <div className="mb-1 flex flex-wrap items-center gap-2">
        <h2 className="text-lg font-semibold text-slate-900 dark:text-slate-100">Histórico</h2>
        <span className="rounded-full bg-gradient-to-r from-brand-from to-brand-to px-2.5 py-0.5 text-xs font-medium text-white">
          Beta
        </span>
        <span className="text-xs text-slate-500 dark:text-slate-400">Em desenvolvimento</span>
      </div>
      <p className="mb-4 text-sm text-slate-600 dark:text-slate-300">
        Painel cronológico das últimas operações realizadas no navegador.
      </p>

      <div className="mb-3 flex flex-wrap items-center gap-2">
        <input
          id="historico_search"
          type="search"
          placeholder="Filtrar por texto"
          aria-label="Filtrar histórico"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="min-w-[220px] flex-1 rounded-lg border border-black/10 bg-white px-3 py-1.5 text-sm text-slate-900 outline-none focus:border-accent focus:ring-2 focus:ring-accent/25 dark:border-white/10 dark:bg-surface-dark-alt dark:text-slate-100"
        />
        <button
          id="historico_export_csv"
          type="button"
          title="Baixar itens visíveis em CSV"
          onClick={exportCsv}
          className="rounded-lg border border-accent px-3 py-1.5 text-sm font-medium text-accent hover:bg-accent/10 dark:border-accent-dark dark:text-accent-dark"
        >
          Exportar CSV
        </button>
        <button
          id="historico_export_json"
          type="button"
          title="Baixar itens visíveis em JSON"
          onClick={exportJson}
          className="rounded-lg border border-black/15 px-3 py-1.5 text-sm font-medium text-slate-700 hover:bg-black/5 dark:border-white/15 dark:text-slate-200 dark:hover:bg-white/5"
        >
          Exportar JSON
        </button>
        <button
          id="clearHistoryBtn"
          type="button"
          aria-label="Limpar histórico"
          title="Remover todos os itens"
          onClick={handleClear}
          className="ml-auto flex items-center gap-1.5 rounded-lg bg-gradient-to-r from-[#ff416c] to-[#ff4b2b] px-3 py-1.5 text-sm font-medium text-white hover:brightness-105"
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth={2}
            aria-hidden="true"
          >
            <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
          </svg>
          Limpar
        </button>
      </div>

      <div id="historico_summary" className="mb-2 text-xs text-slate-500 dark:text-slate-400" aria-live="polite">
        {`Itens visíveis: ${filtered.length} (local: ${state.localCount} | API: ${state.serverCount})`}
      </div>

      <div className="overflow-x-auto rounded-lg border border-black/10 dark:border-white/10" id="historico_table_wrap">
        <table id="historico_table" aria-describedby="historico_summary" className="w-full text-sm">
          <thead className="bg-black/[.03] dark:bg-white/[.04]">
            <tr>
              <th scope="col" className="whitespace-nowrap px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">
                Data/Hora
              </th>
              <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">
                Ação
              </th>
            </tr>
          </thead>
          <tbody id="historico_tbody">
            {filtered.map((item, i) => (
              <tr key={`${item.ts}-${i}`} className="border-t border-black/5 dark:border-white/5">
                <td className="whitespace-nowrap px-3 py-1.5 text-slate-500 dark:text-slate-400">{formatTs(item.ts)}</td>
                <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">
                  {item.text}
                  {item.source === "server" && (
                    <span className="ml-1.5 rounded bg-black/5 px-1.5 py-0.5 text-[11px] text-slate-600 dark:bg-white/10 dark:text-slate-300">
                      API
                    </span>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
