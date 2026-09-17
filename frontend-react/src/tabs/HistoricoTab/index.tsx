import { useEffect, useMemo, useState } from "react";
import { formatTs } from "../../history/formatTs";
import { clearAll, getSnapshot, refreshFromServer, subscribe, type HistoryState } from "../../history/historyStore";
import { downloadFile } from "../../lib/downloadFile";
import { pushToast } from "../../toast/toastStore";
import { Badge } from "../../ui/Badge";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { EmptyState } from "../../ui/EmptyState";
import { IconClock } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { TextInput } from "../../ui/TextInput";

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
    <div>
      <PageHeader
        title="Histórico"
        description="Painel cronológico das últimas operações realizadas no navegador."
        icon={<IconClock className="h-5 w-5" />}
        actions={<Badge tone="brand">Beta</Badge>}
      />

      <Card>
        <div className="mb-3 flex flex-wrap items-center gap-2">
          <TextInput
            id="historico_search"
            type="search"
            placeholder="Filtrar por texto"
            aria-label="Filtrar histórico"
            className="min-w-[220px] flex-1"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          <Button
            id="historico_export_csv"
            variant="outline"
            size="sm"
            title="Baixar itens visíveis em CSV"
            onClick={exportCsv}
          >
            Exportar CSV
          </Button>
          <Button
            id="historico_export_json"
            variant="secondary"
            size="sm"
            title="Baixar itens visíveis em JSON"
            onClick={exportJson}
          >
            Exportar JSON
          </Button>
          <Button
            id="clearHistoryBtn"
            variant="danger"
            size="sm"
            aria-label="Limpar histórico"
            title="Remover todos os itens"
            className="ml-auto"
            onClick={handleClear}
          >
            Limpar
          </Button>
        </div>

        <div id="historico_summary" className="mb-2 text-xs text-text-subtle" aria-live="polite">
          {`Itens visíveis: ${filtered.length} (local: ${state.localCount} | API: ${state.serverCount})`}
        </div>

        <div className="overflow-x-auto rounded-surface border border-border" id="historico_table_wrap">
          <table id="historico_table" aria-describedby="historico_summary" className="w-full text-sm">
            <thead className="bg-surface-sunken">
              <tr>
                <th scope="col" className="whitespace-nowrap px-3 py-2 text-left font-medium text-text-muted">
                  Data/Hora
                </th>
                <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">
                  Ação
                </th>
              </tr>
            </thead>
            <tbody id="historico_tbody">
              {filtered.map((item, i) => (
                <tr key={`${item.ts}-${i}`} className="border-t border-border">
                  <td className="whitespace-nowrap px-3 py-1.5 text-text-muted">{formatTs(item.ts)}</td>
                  <td className="px-3 py-1.5 text-text">
                    {item.text}
                    {item.source === "server" && (
                      <span className="ml-1.5 rounded bg-surface-sunken px-1.5 py-0.5 text-xs text-text-muted">API</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {filtered.length === 0 && (
          <EmptyState
            icon={<IconClock className="h-5 w-5" />}
            title={search ? "Nenhum item corresponde ao filtro" : "Sem histórico ainda"}
            description={
              search
                ? "Ajuste o texto do filtro para ver mais resultados."
                : "As operações que você rodar aparecem aqui automaticamente."
            }
          />
        )}
      </Card>
    </div>
  );
}
