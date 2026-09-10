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
    <div className="tw:rounded-xl tw:border tw:border-black/10 tw:bg-surface tw:p-4 tw:shadow-sm tw:dark:border-white/10 tw:dark:bg-surface-dark tw:sm:p-6">
      <div className="tw:mb-1 tw:flex tw:flex-wrap tw:items-center tw:gap-2">
        <h2 className="tw:text-lg tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">Histórico</h2>
        <span className="tw:rounded-full tw:bg-gradient-to-r tw:from-brand-from tw:to-brand-to tw:px-2.5 tw:py-0.5 tw:text-xs tw:font-medium tw:text-white">
          Beta
        </span>
        <span className="tw:text-xs tw:text-slate-500 tw:dark:text-slate-400">Em desenvolvimento</span>
      </div>
      <p className="tw:mb-4 tw:text-sm tw:text-slate-600 tw:dark:text-slate-300">
        Painel cronológico das últimas operações realizadas no navegador.
      </p>

      <div className="tw:mb-3 tw:flex tw:flex-wrap tw:items-center tw:gap-2">
        <input
          id="historico_search"
          type="search"
          placeholder="Filtrar por texto"
          aria-label="Filtrar histórico"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="tw:min-w-[220px] tw:flex-1 tw:rounded-lg tw:border tw:border-black/10 tw:bg-white tw:px-3 tw:py-1.5 tw:text-sm tw:text-slate-900 tw:outline-none tw:focus:border-accent tw:focus:ring-2 tw:focus:ring-accent/25 tw:dark:border-white/10 tw:dark:bg-surface-dark-alt tw:dark:text-slate-100"
        />
        <button
          id="historico_export_csv"
          type="button"
          title="Baixar itens visíveis em CSV"
          onClick={exportCsv}
          className="tw:rounded-lg tw:border tw:border-accent tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-accent tw:hover:bg-accent/10 tw:dark:border-accent-dark tw:dark:text-accent-dark"
        >
          Exportar CSV
        </button>
        <button
          id="historico_export_json"
          type="button"
          title="Baixar itens visíveis em JSON"
          onClick={exportJson}
          className="tw:rounded-lg tw:border tw:border-black/15 tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-slate-700 tw:hover:bg-black/5 tw:dark:border-white/15 tw:dark:text-slate-200 tw:dark:hover:bg-white/5"
        >
          Exportar JSON
        </button>
        <button
          id="clearHistoryBtn"
          type="button"
          aria-label="Limpar histórico"
          title="Remover todos os itens"
          onClick={handleClear}
          className="tw:ml-auto tw:flex tw:items-center tw:gap-1.5 tw:rounded-lg tw:bg-gradient-to-r tw:from-[#ff416c] tw:to-[#ff4b2b] tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-white tw:hover:brightness-105"
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

      <div id="historico_summary" className="tw:mb-2 tw:text-xs tw:text-slate-500 tw:dark:text-slate-400" aria-live="polite">
        {`Itens visíveis: ${filtered.length} (local: ${state.localCount} | API: ${state.serverCount})`}
      </div>

      <div className="tw:overflow-x-auto tw:rounded-lg tw:border tw:border-black/10 tw:dark:border-white/10" id="historico_table_wrap">
        <table id="historico_table" aria-describedby="historico_summary" className="tw:w-full tw:text-sm">
          <thead className="tw:bg-black/[.03] tw:dark:bg-white/[.04]">
            <tr>
              <th scope="col" className="tw:whitespace-nowrap tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">
                Data/Hora
              </th>
              <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">
                Ação
              </th>
            </tr>
          </thead>
          <tbody id="historico_tbody">
            {filtered.map((item, i) => (
              <tr key={`${item.ts}-${i}`} className="tw:border-t tw:border-black/5 tw:dark:border-white/5">
                <td className="tw:whitespace-nowrap tw:px-3 tw:py-1.5 tw:text-slate-500 tw:dark:text-slate-400">{formatTs(item.ts)}</td>
                <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">
                  {item.text}
                  {item.source === "server" && (
                    <span className="tw:ml-1.5 tw:rounded tw:bg-black/5 tw:px-1.5 tw:py-0.5 tw:text-[11px] tw:text-slate-600 tw:dark:bg-white/10 tw:dark:text-slate-300">
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
