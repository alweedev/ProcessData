(function () {
  const HISTORY_KEY = "history_v2";
  const API_URL = "/api/history";

  const refs = {
    tbody: document.getElementById("historico_tbody"),
    search: document.getElementById("historico_search"),
    summary: document.getElementById("historico_summary"),
    exportCsvBtn: document.getElementById("historico_export_csv"),
    exportJsonBtn: document.getElementById("historico_export_json"),
    clearBtn: document.getElementById("clearHistoryBtn"),
  };

  let filteredHistory = [];
  let serverHistory = [];

  const escapeHtml = (value) =>
    String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;");

  function toast(message, type) {
    if (typeof window.showToast === "function") {
      window.showToast(message, type);
    }
  }

  function migrateOldHistory() {
    const current = localStorage.getItem(HISTORY_KEY);
    if (current) return;
    const legacy = JSON.parse(localStorage.getItem("history") || "[]");
    if (Array.isArray(legacy) && legacy.length) {
      const now = Date.now();
      const migrated = legacy.map((text, index) => ({
        ts: now - (legacy.length - index) * 1000,
        text: String(text),
        source: "local",
      }));
      localStorage.setItem(HISTORY_KEY, JSON.stringify(migrated.slice(-200)));
      try {
        localStorage.removeItem("history");
      } catch (_) {
        // ignore
      }
    }
  }

  function getLocalHistory() {
    const arr = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
    return Array.isArray(arr) ? arr : [];
  }

  function setLocalHistory(arr) {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(arr.slice(-500)));
  }

  async function fetchServerHistory() {
    try {
      const res = await fetch(API_URL + "?limit=300");
      if (!res.ok) return [];
      const data = await res.json();
      if (!data || !Array.isArray(data.items)) return [];
      return data.items.map((item) => ({
        ts: Date.parse(item.timestamp || "") || Date.now(),
        text: `${item.event_type || "evento"}: ${item.status || "status"}`,
        details: item.details || {},
        source: "server",
      }));
    } catch (_) {
      return [];
    }
  }

  function normalizeEntry(entry) {
    if (!entry) return { ts: Date.now(), text: "", source: "local" };
    return {
      ts: Number(entry.ts) || Date.now(),
      text: String(entry.text || ""),
      source: entry.source || "local",
      details: entry.details || {},
    };
  }

  function mergeHistory() {
    const local = getLocalHistory().map(normalizeEntry);
    const all = local.concat(serverHistory.map(normalizeEntry));
    all.sort((a, b) => b.ts - a.ts);
    return all.slice(0, 500);
  }

  function formatTs(ts) {
    try {
      const d = new Date(ts);
      const pad = (n) => String(n).padStart(2, "0");
      return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
    } catch (_) {
      return "-";
    }
  }

  function renderHistory() {
    const all = mergeHistory();
    const term = (refs.search?.value || "").trim().toLowerCase();
    filteredHistory = term
      ? all.filter((item) => item.text.toLowerCase().includes(term))
      : all.slice();

    if (refs.tbody) {
      refs.tbody.innerHTML = filteredHistory
        .map(
          (item) => `<tr>
            <td style="white-space:nowrap;">${formatTs(item.ts)}</td>
            <td>${escapeHtml(item.text)}${item.source === "server" ? " <span class=\"badge bg-light text-dark\">API</span>" : ""}</td>
          </tr>`
        )
        .join("");
    }

    if (refs.summary) {
      const totalLocal = getLocalHistory().length;
      const totalServer = serverHistory.length;
      refs.summary.textContent = `Itens visíveis: ${filteredHistory.length} (local: ${totalLocal} | API: ${totalServer})`;
    }
  }

  function exportCsv() {
    if (!filteredHistory.length) {
      toast("Nada para exportar.", "warning");
      return;
    }
    const header = "data_hora,origem,acao";
    const rows = filteredHistory.map((h) => {
      const action = String(h.text || "").replace(/"/g, '""');
      return `${formatTs(h.ts).replace(/,/g, " ")},${h.source || "local"},"${action}"`;
    });
    const csv = [header].concat(rows).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `historico_${Date.now()}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
    toast("CSV exportado.", "success");
  }

  function exportJson() {
    if (!filteredHistory.length) {
      toast("Nada para exportar.", "warning");
      return;
    }
    const blob = new Blob([JSON.stringify(filteredHistory, null, 2)], {
      type: "application/json",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `historico_${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
    toast("JSON exportado.", "success");
  }

  async function clearHistory() {
    setLocalHistory([]);
    try {
      await fetch(API_URL, { method: "DELETE" });
    } catch (_) {
      // ignore
    }
    serverHistory = [];
    renderHistory();
    toast("Histórico limpo!", "success");
  }

  function addToHistory(action) {
    const history = getLocalHistory();
    history.push({ ts: Date.now(), text: String(action), source: "local" });
    setLocalHistory(history);
    renderHistory();
  }

  async function refresh() {
    serverHistory = await fetchServerHistory();
    renderHistory();
  }

  function bind() {
    refs.search?.addEventListener("input", renderHistory);
    refs.exportCsvBtn?.addEventListener("click", exportCsv);
    refs.exportJsonBtn?.addEventListener("click", exportJson);
    refs.clearBtn?.addEventListener("click", clearHistory);
  }

  function init() {
    migrateOldHistory();
    bind();
    refresh();
  }

  window.addToHistory = addToHistory;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
