document.addEventListener("DOMContentLoaded", function () {
  // ===== UTILITIES & GLOBAL HELPERS =====
  const escapeHtml = (value) =>
    String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  const debounce = (fn, delay = 200) => {
    let timeout;
    return (...args) => {
      clearTimeout(timeout);
      timeout = setTimeout(() => fn.apply(null, args), delay);
    };
  };

  function setStatus(elem, html) {
    if (!elem) return;
    elem.innerHTML = html || "";
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function applyRasterInvertToUploadZones() {
    try {
      const invert = document.body.classList.contains("dark");
      document.querySelectorAll(".upload_dropZone img").forEach((img) => {
        img.classList.toggle("icon-invert-dark", invert);
      });
    } catch (error) {
      console.debug("applyRasterInvertToUploadZones", error);
    }
  }
  // Chamado pelo ThemeToggle (React, Fase 0 da migração) após alternar o tema.
  try { window.__applyRasterInvertToUploadZones = applyRasterInvertToUploadZones; } catch(_) {}

  async function ensureXlsx() {
    if (window.XLSX) return;
    const script = document.createElement("script");
    script.src =
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(script);
    await new Promise((resolve, reject) => {
      script.onload = resolve;
      script.onerror = reject;
    });
  }

  // postFiles/postFormDataJson: removidos (Fase 3 da migração) - só eram
  // usados pelos fluxos de Cadastro e Inativação legados, ambos já
  // portados pro React (src/lib/api.ts tem os equivalentes tipados).

  // showToast: implementado pelo React (src/toast/legacyBridge.ts, Fase 0 da
  // migração), que expõe window.showToast. Continua chamável como `showToast(...)`
  // aqui porque este script roda em escopo clássico (não-module) e todo o
  // código abaixo só executa em DOMContentLoaded, depois que o bundle React
  // (carregado por último, também deferred) já instalou a ponte.

  // ===== Upload banner animations helpers =====
  function setBannerContent(elem, html) {
    if (!elem) return;
    elem.classList.remove('upload-animate-out');
    if (html) {
      elem.innerHTML = html;
      // Reiniciar animação se já estava visível
      elem.classList.remove('upload-animate-in');
      void elem.offsetWidth; // reflow
      elem.classList.add('upload-animate-in');
      const onEnd = (e) => { if (e.animationName === 'bannerIn') { elem.classList.remove('upload-animate-in'); elem.removeEventListener('animationend', onEnd); } };
      elem.addEventListener('animationend', onEnd);
    } else {
      elem.innerHTML = '';
    }
  }
  function clearBanner(elem) {
    if (!elem || !elem.innerHTML) return;
    elem.classList.remove('upload-animate-in');
    elem.classList.add('upload-animate-out');
    const onEnd = (e) => {
      if (e.animationName === 'bannerOut') {
        elem.innerHTML = '';
        elem.classList.remove('upload-animate-out');
        elem.removeEventListener('animationend', onEnd);
      }
    };
    elem.addEventListener('animationend', onEnd);
  }

  function renderPreviewStatsHtml(stats) {
    if (!stats || typeof stats !== "object") return "";
    const parts = [];
    for (const [key, value] of Object.entries(stats)) {
      const normalized = key.toLowerCase();
      const label = normalized.includes("cpf")
        ? "CPF"
        : normalized.includes("exact")
        ? "Exato"
        : normalized.includes("token")
        ? "Token"
        : normalized.includes("fuzzy")
        ? "Fuzzy"
        : key;
      try {
        if (
          value === null ||
          ["string", "number", "boolean"].includes(typeof value)
        ) {
          parts.push(
            `<span class="badge bg-info text-dark me-1 preview-badge">${escapeHtml(
              label
            )}: ${escapeHtml(value)}</span>`
          );
        } else if (Array.isArray(value)) {
          parts.push(
            `<div class="mt-2"><strong>${escapeHtml(
              label
            )}:</strong> <span class="small text-muted">array[${
              value.length
            }]</span></div>`
          );
        } else if (typeof value === "object") {
          const keys = Object.keys(value || {})
            .slice(0, 5)
            .join(", ");
          parts.push(
            `<div class="mt-2"><strong>${escapeHtml(
              label
            )}:</strong> <span class="small text-muted">object{${escapeHtml(
              keys
            )}${
              Object.keys(value || {}).length > 5 ? ", ..." : ""
            }}</span></div>`
          );
        } else {
          parts.push(
            `<div class="mt-2"><strong>${escapeHtml(
              label
            )}:</strong> ${escapeHtml(String(value))}</div>`
          );
        }
      } catch (error) {
        parts.push(
          `<div class="mt-2"><strong>${escapeHtml(
            label
          )}:</strong> ${escapeHtml(String(value))}</div>`
        );
      }
    }
    return `<div class="mt-3"><h6>Estatísticas</h6><div>${parts.join(
      ""
    )}</div></div>`;
  }

  // Tema, animações e status de API: migrados pro React na Fase 0
  // (src/chrome/{ThemeToggle,MotionToggle,ApiStatusBadge}.tsx). O aplicativo
  // React monta em #appChromeControls e chama
  // window.__applyRasterInvertToUploadZones() após alternar o tema.
  applyRasterInvertToUploadZones();

  class AnaliseWorkspace {
    constructor(root) {
      this.root = root;
      this.refs = {
        dropTarget: document.getElementById("analise_dropTarget"),
        pickFileBtn: document.getElementById("analise_pickFile"),
        fileInput: document.getElementById("analise_uploadInput"),
        uploadFeedback: document.getElementById("analise_uploadFeedback"),
        progress: document.getElementById("analise_progress"),
        progressBar: document.querySelector("#analise_progress .analise-progress-bar"),
        summaryStatus: document.getElementById("analise_summaryStatus"),
        metricTotal: document.getElementById("analise_metricTotal"),
        metricColumns: document.getElementById("analise_metricColumns"),
        metricUpdated: document.getElementById("analise_metricUpdated"),
        pinSummary: document.getElementById("analise_pinSummary"),
        summaryCard: document.getElementById("analise_summaryCard"),
        prefWrap: document.getElementById("analise_pref_wrap"),
        prefNowrap: document.getElementById("analise_pref_nowrap"),
        prefDensity: document.getElementById("analise_pref_density"),
        toolbar: document.getElementById("analise_toolbar"),
        searchInput: document.getElementById("analise_searchInput"),
        toggleFull: document.getElementById("analise_toggleFull"),
        exportCsv: document.getElementById("analise_exportCsv"),
        chipTray: document.getElementById("analise_chipTray"),
        emptyState: document.getElementById("analise_emptyState"),
        tableViewport: document.getElementById("analise_tableViewport"),
        tableHead: document.getElementById("analise_tableHead"),
        tableBody: document.getElementById("analise_tableBody"),
        footer: document.getElementById("analise_footer"),
        footerText: document.getElementById("analise_footerText"),
        scrollTop: document.getElementById("analise_scrollTop"),
        debug: document.getElementById("analise_debug"),
        resetBtn: document.getElementById("analise_resetBtn"),
      };

      this.state = {
        headers: [],
        rows: [],
        activeRows: [],
        renderCursor: 0,
        searchTerm: "",
        sheetName: "",
        fileName: "",
        searchLimited: false,
      };
      this.rowChunkSize = 220;
      this.maxSearchResults = 10000;
      this.prefs = this.loadPreferences();
      this.handleSearchDebounced = debounce((value) =>
        this.handleSearch(value)
      );
      this.handleScroll = this.handleScroll.bind(this);
      this.handleEsc = this.handleEsc.bind(this);
      this.isFullscreen = false;

      this.bindEvents();
      this.applyPreferences();
      this.syncPrefButtons();
      this.updateChips();
    }

    loadPreferences() {
      try {
        const stored = JSON.parse(localStorage.getItem("analise_prefs") || "{}");
        return {
          wrapMode: stored.wrapMode || "wrap",
          compact: !!stored.compact,
          summaryPinned: !!stored.summaryPinned,
        };
      } catch (error) {
        return { wrapMode: "wrap", compact: false, summaryPinned: false };
      }
    }

    persistPreferences() {
      localStorage.setItem("analise_prefs", JSON.stringify(this.prefs));
    }

    bindEvents() {
      const {
        dropTarget,
        pickFileBtn,
        fileInput,
        prefWrap,
        prefNowrap,
        prefDensity,
        pinSummary,
        searchInput,
        toggleFull,
        exportCsv,
        scrollTop,
        tableViewport,
        resetBtn,
      } = this.refs;

      if (dropTarget) {
        ["dragenter", "dragover"].forEach((eventName) =>
          dropTarget.addEventListener(eventName, (event) => {
            event.preventDefault();
            dropTarget.classList.add("dragover");
          })
        );
        ["dragleave", "drop"].forEach((eventName) =>
          dropTarget.addEventListener(eventName, (event) => {
            event.preventDefault();
            dropTarget.classList.remove("dragover");
          })
        );
        dropTarget.addEventListener("keydown", (event) => {
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            this.refs.fileInput?.click();
          }
        });
        dropTarget.addEventListener("drop", (event) => {
          const files = event.dataTransfer?.files;
          if (files?.length) {
            this.handleFile(files[0]);
          }
        });
      }

      if (pickFileBtn) {
        pickFileBtn.addEventListener("click", () => this.refs.fileInput?.click());
      }

      if (fileInput) {
        fileInput.addEventListener("change", (event) => {
          const target = event.target;
          const file = target.files?.[0];
          if (file) {
            this.handleFile(file);
          }
          target.value = "";
        });
      }

      prefWrap?.addEventListener("click", () => this.setWrapMode("wrap"));
      prefNowrap?.addEventListener("click", () => this.setWrapMode("nowrap"));
      prefDensity?.addEventListener("click", () => this.toggleDensity());
      pinSummary?.addEventListener("click", () => this.toggleSummaryPin());
      searchInput?.addEventListener("input", (event) =>
        this.handleSearchDebounced(event.target.value)
      );
      toggleFull?.addEventListener("click", () => this.toggleFullscreen());
      exportCsv?.addEventListener("click", () => this.exportCsv());
      scrollTop?.addEventListener("click", () => this.scrollToTop());
      tableViewport?.addEventListener("scroll", this.handleScroll);
      resetBtn?.addEventListener("click", () => this.reset(true));
    }

    handleFile(file) {
      if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showToast("Envie um arquivo Excel (.xlsx ou .xls).", "danger");
        return;
      }
      this.processFile(file);
    }

    async processFile(file) {
      try {
        if (this.refs.uploadFeedback) {
          setBannerContent(this.refs.uploadFeedback, `
            <span class="upload-check" aria-hidden="true">
              <i class="bi bi-check-lg text-success" style="font-size:1.15rem; line-height:1; display:inline-block;"></i>
            </span>
            ${escapeHtml(file.name)}`);
        }
        this.setProgress(8, `Preparando ${file.name}`);
        await ensureXlsx();
        this.setProgress(35, "Lendo planilha...");
        const buffer = await file.arrayBuffer();
        this.setProgress(65, "Normalizando dados...");
        const payload = this.parseBuffer(buffer);
        this.setProgress(92, "Renderizando");
        this.afterDataLoaded(payload, file);
        this.setProgress(100, "Concluído");
        setTimeout(() => this.resetProgress(), 500);
        showToast("Ficha carregada com sucesso!", "success");
        window.addToHistory?.(
          `Análise: ${file.name} - ${new Date().toLocaleString("pt-BR")}`
        );
        try {
          window.__signalUploadSuccess?.(file.name);
        } catch (error) {
          console.debug("signal upload", error);
        }
      } catch (error) {
        this.resetProgress();
        this.handleError(error);
      }
    }

    parseBuffer(buffer) {
      const workbook = XLSX.read(buffer, { type: "array", dense: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        blankrows: false,
      });
      if (!matrix.length) {
        throw new Error("Planilha vazia");
      }
      const headers = matrix[0].map((value, index) => {
        const label = String(value ?? "").trim();
        return label || `Coluna ${index + 1}`;
      });
      const rows = matrix.slice(1).map((row) => {
        const entry = {};
        headers.forEach((header, index) => {
          const cell = row[index];
          entry[header] = cell === undefined || cell === null ? "" : String(cell).trim();
        });
        return entry;
      }).filter((row) => headers.some((header) => row[header] !== ""));
      return { headers, rows, sheetName };
    }

    afterDataLoaded(payload, file) {
      const { headers, rows, sheetName } = payload;
      this.state.headers = headers;
      this.state.rows = rows;
      this.state.activeRows = rows.slice();
      this.state.sheetName = sheetName;
      this.state.fileName = file.name;
      this.state.renderCursor = 0;
      this.state.searchTerm = "";
      this.state.searchLimited = false;

      this.updateSummary(rows.length, headers.length, file.name, sheetName);
      this.renderHeaders();
      this.renderRows();
      this.toggleEmptyState(false);
      this.setToolbarVisible(true);
      this.updateChips();
      this.refs.searchInput && (this.refs.searchInput.value = "");
    }

    updateSummary(totalRows, totalCols, fileName, sheetName) {
      const now = new Date();
      this.refs.metricTotal.textContent = totalRows.toLocaleString("pt-BR");
      this.refs.metricColumns.textContent = totalCols.toString();
      this.refs.metricUpdated.textContent = now.toLocaleTimeString("pt-BR", {
        hour: "2-digit",
        minute: "2-digit",
      });
      this.refs.summaryStatus.textContent = `${fileName} • ${sheetName} • ${totalRows.toLocaleString(
        "pt-BR"
      )} linhas`;
    }

    renderHeaders() {
      if (!this.refs.tableHead) return;
      if (!this.state.headers.length) {
        this.refs.tableHead.innerHTML = "";
        return;
      }
      const html =
        "<tr>" +
        this.state.headers
          .map((header) => `<th scope="col">${escapeHtml(header)}</th>`)
          .join("") +
        "</tr>";
      this.refs.tableHead.innerHTML = html;
    }

    renderRows() {
      const rows = this.state.activeRows;
      this.refs.tableBody.innerHTML = "";
      this.state.renderCursor = 0;
      if (!rows.length) {
        if (this.refs.footerText) {
          this.refs.footerText.textContent = "Sem linhas correspondentes";
        }
        if (this.refs.footer) this.refs.footer.hidden = false;
        this.refs.tableViewport?.setAttribute("hidden", "true");
        this.toggleEmptyState(true);
        return;
      }
      this.refs.tableViewport?.removeAttribute("hidden");
      if (this.refs.footer) this.refs.footer.hidden = false;
      this.toggleEmptyState(false);
      this.appendNextChunk();
    }

    appendNextChunk() {
      const rows = this.state.activeRows;
      if (this.state.renderCursor >= rows.length) return;
      const start = this.state.renderCursor;
      const end = Math.min(start + this.rowChunkSize, rows.length);
      const slice = rows.slice(start, end);
      const html = slice
        .map((row) =>
          "<tr>" +
            this.state.headers
              .map((header) => `<td>${escapeHtml(row[header] ?? "")}</td>`)
              .join("") +
            "</tr>"
        )
        .join("");
      this.refs.tableBody.insertAdjacentHTML("beforeend", html);
      this.state.renderCursor = end;
      this.updateFooter();
    }

    updateFooter() {
      if (!this.refs.footerText) return;
      const rendered = this.state.renderCursor;
      const total = this.state.activeRows.length;
      const parts = [`${rendered.toLocaleString("pt-BR")} / ${total.toLocaleString("pt-BR")} linhas visíveis`];
      if (this.state.searchTerm) {
        parts.push(`Filtro ativo: "${escapeHtml(this.state.searchTerm)}"`);
        if (this.state.searchLimited) {
          parts.push(`Limite de ${this.maxSearchResults.toLocaleString("pt-BR")} resultados`);
        }
      }
      this.refs.footerText.innerHTML = parts.join(" • ");
    }

    handleScroll(event) {
      if (this.state.renderCursor >= this.state.activeRows.length) return;
      const viewport = event.target;
      const nearBottom =
        viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 120;
      if (nearBottom && !this.state.loadingChunk) {
        this.state.loadingChunk = true;
        requestAnimationFrame(() => {
          this.appendNextChunk();
          this.state.loadingChunk = false;
        });
      }
    }

    handleSearch(value) {
      const term = (value || "").trim().toLowerCase();
      this.state.searchTerm = term;
      this.state.searchLimited = false;
      if (!term) {
        this.state.activeRows = this.state.rows.slice();
        this.renderRows();
        this.updateChips();
        return;
      }
      const filtered = [];
      for (const row of this.state.rows) {
        const match = this.state.headers.some((header) =>
          String(row[header] || "").toLowerCase().includes(term)
        );
        if (match) {
          filtered.push(row);
          if (filtered.length >= this.maxSearchResults) {
            this.state.searchLimited = true;
            break;
          }
        }
      }
      this.state.activeRows = filtered;
      this.renderRows();
      this.updateChips();
      if (!filtered.length) {
        this.refs.footerText.textContent = "Nenhum resultado encontrado";
      }
    }

    updateChips() {
      if (!this.refs.chipTray) return;
      const chips = [];
      if (this.state.searchTerm) {
        chips.push({
          label: `Filtro: ${this.state.searchTerm}`,
          action: () => {
            this.state.searchTerm = "";
            this.refs.searchInput && (this.refs.searchInput.value = "");
            this.state.activeRows = this.state.rows.slice();
            this.renderRows();
            this.updateChips();
          },
        });
      }
      if (this.prefs.wrapMode === "nowrap") {
        chips.push({ label: "Texto contínuo", action: () => this.setWrapMode("wrap") });
      }
      if (this.prefs.compact) {
        chips.push({ label: "Modo compacto", action: () => this.toggleDensity() });
      }
      if (!chips.length) {
        this.refs.chipTray.innerHTML = '<span class="text-muted small">Preferências rápidas aparecem aqui.</span>';
        return;
      }
      this.refs.chipTray.innerHTML = chips
        .map(
          (chip, index) =>
            `<button type="button" class="analise-chip" data-chip-index="${index}">${escapeHtml(
              chip.label
            )}</button>`
        )
        .join("");
      this.refs.chipTray.querySelectorAll("[data-chip-index]").forEach((btn) => {
        const idx = Number(btn.getAttribute("data-chip-index"));
        btn.addEventListener("click", () => chips[idx].action());
      });
    }

    setWrapMode(mode) {
      if (this.prefs.wrapMode === mode) return;
      this.prefs.wrapMode = mode;
      this.persistPreferences();
      this.applyPreferences();
      this.syncPrefButtons();
      this.updateChips();
    }

    toggleDensity() {
      this.prefs.compact = !this.prefs.compact;
      this.persistPreferences();
      this.applyPreferences();
      this.updateChips();
    }

    toggleSummaryPin() {
      this.prefs.summaryPinned = !this.prefs.summaryPinned;
      this.persistPreferences();
      document.body.classList.toggle(
        "analise-summary-pinned",
        this.prefs.summaryPinned
      );
      if (this.refs.pinSummary) {
        this.refs.pinSummary.setAttribute(
          "aria-pressed",
          this.prefs.summaryPinned ? "true" : "false"
        );
      }
    }

    applyPreferences() {
      document.body.classList.toggle(
        "analise-wrap",
        this.prefs.wrapMode === "wrap"
      );
      document.body.classList.toggle(
        "analise-nowrap",
        this.prefs.wrapMode === "nowrap"
      );
      document.body.classList.toggle("compact", this.prefs.compact);
      document.body.classList.toggle(
        "analise-summary-pinned",
        this.prefs.summaryPinned
      );
      if (this.refs.pinSummary) {
        this.refs.pinSummary.setAttribute(
          "aria-pressed",
          this.prefs.summaryPinned ? "true" : "false"
        );
      }
    }

    syncPrefButtons() {
      const { prefWrap, prefNowrap, prefDensity } = this.refs;
      prefWrap?.classList.toggle("active", this.prefs.wrapMode === "wrap");
      prefNowrap?.classList.toggle("active", this.prefs.wrapMode === "nowrap");
      prefDensity?.classList.toggle("active", this.prefs.compact);
    }

    toggleFullscreen() {
      if (!this.refs.tableViewport) return;
      const card = document.getElementById("analise_tableCard");
      if (!card) return;
      this.isFullscreen = !this.isFullscreen;
      card.classList.toggle("analise-fullscreen", this.isFullscreen);
      document.body.classList.toggle("analise-fullscreen-active", this.isFullscreen);
      if (this.refs.toggleFull) {
        this.refs.toggleFull.textContent = this.isFullscreen
          ? "Sair"
          : "Tela cheia";
      }
      if (this.isFullscreen) {
        document.addEventListener("keydown", this.handleEsc);
      } else {
        document.removeEventListener("keydown", this.handleEsc);
      }
    }

    handleEsc(event) {
      if (event.key === "Escape" && this.isFullscreen) {
        this.toggleFullscreen();
      }
    }

    exportCsv() {
      if (!this.state.headers.length || !this.state.activeRows.length) {
        showToast("Nenhum dado para exportar", "info");
        return;
      }
      const limit = 20000;
      const slice = this.state.activeRows.slice(0, limit);
      const csv = this.buildCsv(slice);
      const name = this.state.fileName
        ? this.state.fileName.replace(/\.[^.]+$/, "")
        : "analise";
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      downloadBlob(blob, `${name}_analise.csv`);
      showToast(
        slice.length < limit ? "CSV exportado" : `Exportando ${slice.length.toLocaleString("pt-BR")} linhas (limite)`
      );
    }

    buildCsv(rows) {
      const escapeValue = (value) =>
        `"${String(value ?? "").replace(/"/g, '""')}"`;
      const head = this.state.headers.map(escapeValue).join(",");
      const body = rows.map((row) =>
        this.state.headers.map((header) => escapeValue(row[header])).join(",")
      );
      return [head].concat(body).join("\r\n");
    }

    setProgress(percent, label) {
      if (!this.refs.progress || !this.refs.progressBar) return;
      this.refs.progress.hidden = false;
      this.refs.progressBar.style.width = `${percent}%`;
      this.refs.progressBar.setAttribute("aria-valuenow", percent);
      this.refs.summaryStatus.textContent = label;
    }

    resetProgress() {
      if (!this.refs.progress || !this.refs.progressBar) return;
      this.refs.progress.hidden = true;
      this.refs.progressBar.style.width = "0%";
    }

    setToolbarVisible(state) {
      if (!this.refs.toolbar) return;
      this.refs.toolbar.setAttribute("aria-hidden", state ? "false" : "true");
    }

    toggleEmptyState(show) {
      if (show) {
        this.refs.emptyState?.removeAttribute("hidden");
      } else {
        this.refs.emptyState?.setAttribute("hidden", "true");
      }
    }

    scrollToTop() {
      this.refs.tableViewport?.scrollTo({ top: 0, behavior: "smooth" });
    }

    handleError(error) {
      console.error(error);
      this.refs.debug.textContent = `Erro: ${error.message || error}`;
      showToast(`Falha ao processar: ${error.message || error}`, "danger");
    }

    reset(showNotice) {
      this.state.headers = [];
      this.state.rows = [];
      this.state.activeRows = [];
      this.state.renderCursor = 0;
      this.state.searchTerm = "";
      this.state.fileName = "";
      this.state.sheetName = "";
      this.refs.tableHead.innerHTML = "";
      this.refs.tableBody.innerHTML = "";
      this.refs.summaryStatus.textContent = "Nenhum arquivo processado";
      this.refs.metricTotal.textContent = "—";
      this.refs.metricColumns.textContent = "—";
      this.refs.metricUpdated.textContent = "—";
      this.refs.searchInput && (this.refs.searchInput.value = "");
      this.refs.footer.hidden = true;
      this.refs.tableViewport?.setAttribute("hidden", "true");
      this.setToolbarVisible(false);
      this.toggleEmptyState(true);
      this.updateChips();
      this.resetProgress();
      if (this.refs.uploadFeedback) {
        this.refs.uploadFeedback.textContent = "";
      }
      if (this.isFullscreen) this.toggleFullscreen();
      if (showNotice) {
        showToast("Workspace limpo", "info");
      }
    }
  }

  const analiseRoot = document.querySelector("[data-analise-root]");
  const analiseWorkspace = analiseRoot ? new AnaliseWorkspace(analiseRoot) : null;
  window.Analise = analiseWorkspace || { reset() {} };

  // Persistência de login_choice/fluxo: portada pro React
  // (src/tabs/CadastroTab, Fase 3 da migração).

  // wireCadastroHelp: portado pro React (src/tabs/CadastroTab, Fase 3).
  // wireInativacaoHelp: portado pro React (src/tabs/InativacaoTab/index.tsx,
  // Fase 2 da migração) - o botão #inativacao_help_btn agora é montado por lá.

  // Cadastro: portado pro React (src/tabs/CadastroTab, Fase 3 da migração).
  // O try/catch abaixo continua envolvendo só o código de Estruturas de
  // Aprovação (ainda não migrado).
  try {
    // Estruturas de aprovação: Remover aprovador por CPF
    const aprovUsersInput = document.getElementById('aprovacao_users_file');
    const aprovUsersPickBtn = document.getElementById('aprovacao_users_pick');
    const aprovUsersName = document.getElementById('aprovacao_users_name');
    const aprovUsersDrop = document.getElementById('aprovacao_users_uploadArea');
    const aprovUsersFeedback = document.getElementById('aprovacao_users_feedback');
    const aprovUsersClearBtn = document.getElementById('aprovacao_users_clear_btn');

    const aprovBaseInput = document.getElementById('aprovacao_base_file');
    const aprovBasePickBtn = document.getElementById('aprovacao_base_pick');
    const aprovBaseName = document.getElementById('aprovacao_base_name');
    const aprovBaseDrop = document.getElementById('aprovacao_base_uploadArea');
    const aprovBaseFeedback = document.getElementById('aprovacao_base_feedback');
    const aprovBaseClearBtn = document.getElementById('aprovacao_base_clear_btn');

    const aprovCpfInput = document.getElementById('aprovacao_cpf');
    const aprovPreviewBtn = document.getElementById('aprovacao_preview_btn');
    const aprovRemoveAllBtn = document.getElementById('aprovacao_remove_all_btn');
    const aprovRemoveSelectedBtn = document.getElementById('aprovacao_remove_selected_btn');
    const aprovRemoveSecondLevel = document.getElementById('aprovacao_remove_second_level');

    const aprovStatus = document.getElementById('aprovacao_status');
    const aprovPanel = document.getElementById('aprovacao_preview_panel');
    const aprovApproverName = document.getElementById('aprovacao_approver_name');
    const aprovApproverCpf = document.getElementById('aprovacao_approver_cpf');
    const aprovSummaryEstruturas = document.getElementById('aprovacao_summary_estruturas');
    const aprovSummaryOcorrencias = document.getElementById('aprovacao_summary_ocorrencias');
    const aprovSummaryTipos = document.getElementById('aprovacao_summary_tipos');
    const aprovTableWrap = document.getElementById('aprovacao_table_wrap');
    const aprovTableBody = document.getElementById('aprovacao_table_body');
    const aprovCheckAll = document.getElementById('aprovacao_check_all');
    const aprovSelectedBadge = document.getElementById('aprovacao_selected_badge');
    const aprovSelectedCount = document.getElementById('aprovacao_selected_count');

    let aprovCurrentItems = [];

    function setAprovStatus(message, isError = false) {
      if (!aprovStatus) return;
      aprovStatus.textContent = message || '';
      aprovStatus.classList.toggle('text-danger', !!isError);
      aprovStatus.classList.toggle('text-muted', !isError);
    }

    function updateAprovReadinessHint() {
      if (!aprovStatus) return;
      const hasUsers = !!(aprovUsersInput && aprovUsersInput.files && aprovUsersInput.files[0]);
      const hasBase = !!(aprovBaseInput && aprovBaseInput.files && aprovBaseInput.files[0]);

      if (!hasUsers && !hasBase) {
        setAprovStatus('Envie a base de usuários e a base de aprovação para começar.', false);
        return;
      }
      if (hasUsers && !hasBase) {
        setAprovStatus('Base de usuários selecionada. Agora selecione a base de aprovação.', false);
        return;
      }
      if (!hasUsers && hasBase) {
        setAprovStatus('Base de aprovação selecionada. Agora selecione a base de usuários.', false);
        return;
      }
      setAprovStatus('Bases prontas. Informe o CPF e clique em Verificar.', false);
    }

    function normalizeCpf(value) {
      if (!value) return '';
      return String(value).replace(/\D/g, '');
    }

    function updateAprovSelectedCount() {
      if (!aprovTableBody || !aprovSelectedBadge || !aprovSelectedCount) return;
      const checkboxes = aprovTableBody.querySelectorAll('input.aprov-row-check');
      let selected = 0;
      checkboxes.forEach((cb) => { if (cb.checked) selected += 1; });
      aprovSelectedCount.textContent = String(selected);
      aprovSelectedBadge.style.display = selected > 0 ? 'inline-flex' : 'none';
      if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = selected === 0;
    }

    function renderAprovSummary(data) {
      if (!aprovPanel) return;
      aprovPanel.classList.remove('d-none');

      if (aprovApproverName) aprovApproverName.textContent = data?.approver?.nomeCompleto || '—';
      if (aprovApproverCpf) aprovApproverCpf.textContent = data?.approver?.cpf ? `CPF: ${data.approver.cpf}` : '';
      if (aprovSummaryEstruturas) aprovSummaryEstruturas.textContent = String(data?.summary?.estruturasAfetadas ?? 0);
      if (aprovSummaryOcorrencias) aprovSummaryOcorrencias.textContent = String(data?.summary?.ocorrenciasTotal ?? 0);

      if (aprovSummaryTipos) {
        aprovSummaryTipos.innerHTML = '';
        const map = data?.summary?.porAprovacaoPor || {};
        Object.entries(map).forEach(([k, v]) => {
          const span = document.createElement('span');
          span.className = 'badge bg-light text-body border me-1';
          span.textContent = `${k}: ${v}`;
          aprovSummaryTipos.appendChild(span);
        });
      }
    }

    function renderAprovTable(items) {
      if (!aprovTableBody || !aprovTableWrap) return;
      aprovTableBody.innerHTML = '';

      if (!items || !items.length) {
        aprovTableWrap.classList.add('d-none');
        if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = true;
        if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = true;
        if (aprovCheckAll) aprovCheckAll.checked = false;
        updateAprovSelectedCount();
        return;
      }

      items.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const aprovacaoId = item.aprovacaoId ?? '';
        const aprovacaoPor = item.aprovacaoPor ?? '';
        
        // Destacar linha que ficará sem aprovador
        if (item.ficaraSemAprovador) {
          tr.className = 'table-warning';
        }
        
        const contexto = (() => {
          if (aprovacaoPor === 'VIAJANTE' && item.viajanteNomeCompleto) {
            return item.viajanteNomeCompleto;
          }
          if (aprovacaoPor === 'CCEMPRESA') {
            const cod = item.ccCodigo || '';
            const desc = item.ccDescricao || '';
            if (cod && desc) return `${cod} - ${desc}`;
            return cod || desc || '';
          }
          const cod = item.ccCodigo || '';
          const desc = item.ccDescricao || '';
          if (cod && desc) return `${cod} - ${desc}`;
          if (item.viajanteNomeCompleto) return item.viajanteNomeCompleto;
          return cod || desc || '';
        })();

        const posicoesStr = (item.posicoes || []).length
          ? (item.posicoes || []).join(', ')
          : '—';
        const segundoNivelStr = item.segundoNivel ? 'Sim' : 'Não';
        
        // Ícone de aviso se ficará sem aprovador
        const warningIcon = item.ficaraSemAprovador 
          ? '<i class="bi bi-exclamation-triangle-fill text-warning ms-1" title="Esta estrutura ficará sem aprovadores"></i>' 
          : '';

        tr.innerHTML = `
          <td>
            <input type="checkbox" class="form-check-input aprov-row-check" data-id="${escapeHtml(String(aprovacaoId))}" checked />
          </td>
          <td>${escapeHtml(String(aprovacaoId || ''))}${warningIcon}</td>
          <td>${escapeHtml(String(aprovacaoPor || ''))}</td>
          <td>${escapeHtml(String(item.aprovacao ?? ''))}</td>
          <td>${escapeHtml(String(item.tipo ?? ''))}</td>
          <td>${escapeHtml(String(item.valor ?? ''))}</td>
          <td>${escapeHtml(String(contexto || ''))}</td>
          <td>${escapeHtml(posicoesStr)}</td>
          <td>${escapeHtml(segundoNivelStr)}</td>
        `;
        aprovTableBody.appendChild(tr);
      });

      aprovTableWrap.classList.remove('d-none');
      if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = false;
      if (aprovCheckAll) aprovCheckAll.checked = true;

      // Listener de seleção por linha
      aprovTableBody.querySelectorAll('input.aprov-row-check').forEach((cb) => {
        cb.addEventListener('change', () => {
          updateAprovSelectedCount();
        });
      });

      updateAprovSelectedCount();
    }

    if (aprovCheckAll && aprovTableBody) {
      aprovCheckAll.addEventListener('change', () => {
        const checked = aprovCheckAll.checked;
        aprovTableBody.querySelectorAll('input.aprov-row-check').forEach((cb) => {
          cb.checked = checked;
        });
        updateAprovSelectedCount();
      });
    }

    function handleFileChange(input, labelEl, feedbackWrapper, clearBtn, toastLabel) {
      if (!input) return;
      const file = input.files && input.files[0];
      if (labelEl) {
        labelEl.textContent = file ? file.name : '';
      }
      
      // Novo formato: wrapper com conteúdo + botão de limpar
      if (feedbackWrapper) {
        const contentEl = feedbackWrapper.querySelector('.upload-feedback-content');
        if (file) {
          if (contentEl) {
            contentEl.innerHTML = `<i class="bi bi-check-lg text-success" style="font-size:1.15rem; line-height:1;"></i><span>${escapeHtml(file.name)}</span>`;
          }
          feedbackWrapper.classList.add('has-file');
          if (clearBtn) clearBtn.classList.remove('d-none');
        } else {
          if (contentEl) contentEl.innerHTML = '';
          feedbackWrapper.classList.remove('has-file');
          if (clearBtn) clearBtn.classList.add('d-none');
        }
      }
      
      if (file && toastLabel) {
        showToast(`${toastLabel} carregada com sucesso.`, 'success');
      }
      updateAprovReadinessHint();
    }

    function clearUploadFeedback(feedbackWrapper, clearBtn) {
      if (feedbackWrapper) {
        const contentEl = feedbackWrapper.querySelector('.upload-feedback-content');
        if (contentEl) contentEl.innerHTML = '';
        feedbackWrapper.classList.remove('has-file');
      }
      if (clearBtn) clearBtn.classList.add('d-none');
    }

    if (aprovUsersPickBtn && aprovUsersInput) {
      aprovUsersPickBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        aprovUsersInput.click();
      });
    }
    if (aprovUsersInput) {
      aprovUsersInput.addEventListener('change', () => handleFileChange(aprovUsersInput, aprovUsersName, aprovUsersFeedback, aprovUsersClearBtn, 'Base de usuários'));
    }
    if (aprovUsersDrop && aprovUsersInput) {
      aprovUsersDrop.addEventListener('click', (e) => {
        // Ignorar cliques em botões ou inputs para evitar duplo disparo
        if (e.target.closest('button') || e.target.closest('input')) return;
        aprovUsersInput.click();
      });
    }

    if (aprovBasePickBtn && aprovBaseInput) {
      aprovBasePickBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        aprovBaseInput.click();
      });
    }
    if (aprovBaseInput) {
      aprovBaseInput.addEventListener('change', () => handleFileChange(aprovBaseInput, aprovBaseName, aprovBaseFeedback, aprovBaseClearBtn, 'Base de aprovação'));
    }
    if (aprovBaseDrop && aprovBaseInput) {
      aprovBaseDrop.addEventListener('click', (e) => {
        // Ignorar cliques em botões ou inputs para evitar duplo disparo
        if (e.target.closest('button') || e.target.closest('input')) return;
        aprovBaseInput.click();
      });
    }

    if (aprovUsersClearBtn) {
      aprovUsersClearBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        try { if (aprovUsersInput) aprovUsersInput.value = ''; } catch (_) {}
        if (aprovUsersName) aprovUsersName.textContent = '';
        clearUploadFeedback(aprovUsersFeedback, aprovUsersClearBtn);
        updateAprovReadinessHint();
        showToast('Base de usuários limpa.', 'info');
      });
    }

    if (aprovBaseClearBtn) {
      aprovBaseClearBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        try { if (aprovBaseInput) aprovBaseInput.value = ''; } catch (_) {}
        if (aprovBaseName) aprovBaseName.textContent = '';
        clearUploadFeedback(aprovBaseFeedback, aprovBaseClearBtn);
        updateAprovReadinessHint();
        showToast('Base de aprovação limpa.', 'info');
      });
    }

    // Botão Limpar Tudo: Estruturas de Aprovação
    const aprovClearAllBtn = document.getElementById('aprovacao_clear_all_btn');
    if (aprovClearAllBtn) {
      aprovClearAllBtn.addEventListener('click', async () => {
        // Adiciona animação de execução ao botão
        aprovClearAllBtn.classList.add('executing');
        
        // Anima o painel de preview antes de limpar (se visível)
        if (aprovPanel && !aprovPanel.classList.contains('d-none')) {
          aprovPanel.style.transition = 'opacity 0.35s ease, transform 0.35s ease';
          aprovPanel.style.opacity = '0';
          aprovPanel.style.transform = 'translateY(-10px) scale(0.98)';
          await new Promise(r => setTimeout(r, 350));
        }

        // Limpar inputs de arquivo
        try { if (aprovUsersInput) aprovUsersInput.value = ''; } catch (_) {}
        try { if (aprovBaseInput) aprovBaseInput.value = ''; } catch (_) {}
        
        // Limpar feedbacks visuais
        if (aprovUsersName) aprovUsersName.textContent = '';
        if (aprovBaseName) aprovBaseName.textContent = '';
        clearUploadFeedback(aprovUsersFeedback, aprovUsersClearBtn);
        clearUploadFeedback(aprovBaseFeedback, aprovBaseClearBtn);
        
        // Limpar CPF
        if (aprovCpfInput) aprovCpfInput.value = '';
        
        // Limpar dados internos
        aprovCurrentItems = [];
        
        // Limpar tabela
        if (aprovTableBody) aprovTableBody.innerHTML = '';
        if (aprovTableWrap) aprovTableWrap.classList.add('d-none');
        if (aprovCheckAll) aprovCheckAll.checked = false;
        
        // Limpar sumário
        if (aprovApproverName) aprovApproverName.innerHTML = '&nbsp;';
        if (aprovApproverCpf) aprovApproverCpf.innerHTML = '&nbsp;';
        if (aprovSummaryEstruturas) aprovSummaryEstruturas.textContent = '0';
        if (aprovSummaryOcorrencias) aprovSummaryOcorrencias.textContent = '0';
        if (aprovSummaryTipos) aprovSummaryTipos.innerHTML = '';
        
        // Limpar alertas
        const alertContainer = document.getElementById('aprovacao_alert_container');
        if (alertContainer) alertContainer.innerHTML = '';
        
        // Esconder painel e resetar botões
        if (aprovPanel) {
          aprovPanel.classList.add('d-none');
          aprovPanel.style.opacity = '';
          aprovPanel.style.transform = '';
          aprovPanel.style.transition = '';
        }
        if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = true;
        if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = true;
        if (aprovSelectedBadge) aprovSelectedBadge.style.display = 'none';
        
        // Atualizar status
        updateAprovReadinessHint();
        
        // Remove animação do botão
        setTimeout(() => {
          aprovClearAllBtn.classList.remove('executing');
        }, 600);
        
        showToast('Todos os dados foram limpos.', 'info');
      });
    }

    async function callAprovPreview() {
      if (!aprovUsersInput || !aprovBaseInput) return;

      const usersFile = aprovUsersInput.files && aprovUsersInput.files[0];
      const baseFile = aprovBaseInput.files && aprovBaseInput.files[0];
      const rawCpf = aprovCpfInput ? aprovCpfInput.value : '';
      const cpfDigits = normalizeCpf(rawCpf);

      if (!usersFile || !baseFile) {
        setAprovStatus('Envie a base de usuários e a base de aprovação.', true);
        showToast('Envie a base de usuários e a base de aprovação.', 'danger');
        return;
      }

      if (!cpfDigits || cpfDigits.length !== 11) {
        setAprovStatus('CPF inválido. Informe 11 dígitos.', true);
        showToast('CPF inválido. Informe 11 dígitos.', 'danger');
        return;
      }

      const formData = new FormData();
      formData.append('users_file', usersFile);
      formData.append('base_file', baseFile);
      formData.append('cpf', rawCpf);

      if (aprovPreviewBtn) aprovPreviewBtn.disabled = true;
      if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = true;
      if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = true;
      setAprovStatus('Verificando estruturas...', false);

      try {
        const response = await fetch('/api/aprovacao/remover/preview', { method: 'POST', body: formData });
        const data = await response.json().catch(() => ({}));

        if (!response.ok) {
          const msg = data?.error || 'Erro ao obter preview.';
          setAprovStatus(msg, true);
          showToast(msg, 'danger');
          aprovCurrentItems = [];
          renderAprovTable(aprovCurrentItems);
          return;
        }

        aprovCurrentItems = data.items || [];
        if (!aprovCurrentItems.length) {
          setAprovStatus('Nenhuma estrutura encontrada para o CPF informado.', true);
          showToast('Nenhuma estrutura encontrada para o CPF informado.', 'info');
        } else {
          setAprovStatus(`Encontradas ${aprovCurrentItems.length} estrutura(s) contendo o aprovador.`, false);
        }

        renderAprovSummary(data);
        renderAprovTable(aprovCurrentItems);
      } catch (err) {
        const msg = 'Erro ao contatar o servidor de aprovação.';
        setAprovStatus(msg, true);
        showToast(msg, 'danger');
      } finally {
        if (aprovPreviewBtn) aprovPreviewBtn.disabled = false;
      }
    }

    if (aprovPreviewBtn) {
      aprovPreviewBtn.addEventListener('click', () => {
        callAprovPreview();
      });
    }

    function getSelectedAprovacaoIds() {
      if (!aprovTableBody) return [];
      const ids = [];
      aprovTableBody.querySelectorAll('input.aprov-row-check').forEach((cb) => {
        if (cb.checked) {
          const id = cb.getAttribute('data-id');
          if (id) ids.push(id);
        }
      });
      return ids;
    }

    async function callAprovExport(mode, ignoreEmptyWarning = false) {
      if (!aprovUsersInput || !aprovBaseInput) return;
      const usersFile = aprovUsersInput.files && aprovUsersInput.files[0];
      const baseFile = aprovBaseInput.files && aprovBaseInput.files[0];
      const rawCpf = aprovCpfInput ? aprovCpfInput.value : '';
      const cpfDigits = normalizeCpf(rawCpf);

      if (!aprovCurrentItems || !aprovCurrentItems.length) {
        setAprovStatus('Realize primeiro a verificação do CPF antes de remover.', true);
        showToast('Realize primeiro a verificação do CPF.', 'warning');
        return;
      }

      if (!usersFile || !baseFile) {
        setAprovStatus('Envie novamente a base de usuários e a base de aprovação.', true);
        showToast('Envie novamente a base de usuários e a base de aprovação.', 'danger');
        return;
      }

      if (!cpfDigits || cpfDigits.length !== 11) {
        setAprovStatus('CPF inválido. Informe 11 dígitos.', true);
        showToast('CPF inválido. Informe 11 dígitos.', 'danger');
        return;
      }

      const formData = new FormData();
      formData.append('users_file', usersFile);
      formData.append('base_file', baseFile);
      formData.append('cpf', rawCpf);
      formData.append('mode', mode);

      if (aprovRemoveSecondLevel && aprovRemoveSecondLevel.checked) {
        formData.append('remove_second_level', '1');
      }
      
      if (ignoreEmptyWarning) {
        formData.append('ignore_empty_warning', '1');
      }

      if (mode === 'selected') {
        const selectedIds = getSelectedAprovacaoIds();
        if (!selectedIds.length) {
          setAprovStatus('Nenhuma estrutura selecionada. Marque pelo menos uma linha.', true);
          showToast('Nenhuma estrutura selecionada.', 'warning');
          return;
        }
        selectedIds.forEach((id) => formData.append('selected_aprovacao_ids[]', id));
      }

      if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = true;
      if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = true;
      if (aprovPreviewBtn) aprovPreviewBtn.disabled = true;
      setAprovStatus('Gerando base de aprovação atualizada...', false);

      try {
        const response = await fetch('/api/aprovacao/remover/export', { method: 'POST', body: formData });
        
        if (!response.ok) {
          const errData = await response.json().catch(() => ({}));
          
          // Check if this is a warning about structures without approvers
          if (errData.warning && errData.estruturasSemAprovador && errData.estruturasSemAprovador.length > 0) {
            const estruturas = errData.estruturasSemAprovador;
            const listaHtml = estruturas.map(e => `<li>AprovacaoId: <strong>${escapeHtml(String(e.aprovacaoId))}</strong></li>`).join('');
            
            const result = await Swal.fire({
              title: 'Atenção!',
              html: `
                <p>${escapeHtml(errData.message || '')}</p>
                <p>As seguintes estruturas ficarão <strong>sem nenhum aprovador</strong>:</p>
                <ul class="text-start" style="max-height: 200px; overflow-y: auto;">${listaHtml}</ul>
                <p class="text-danger fw-bold">Deseja continuar mesmo assim?</p>
              `,
              icon: 'warning',
              showCancelButton: true,
              confirmButtonColor: '#d33',
              cancelButtonColor: '#6c757d',
              confirmButtonText: 'Sim, continuar',
              cancelButtonText: 'Cancelar'
            });
            
            if (result.isConfirmed) {
              // Re-enable buttons and retry with ignore_empty_warning flag
              if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = false;
              if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = false;
              if (aprovPreviewBtn) aprovPreviewBtn.disabled = false;
              return callAprovExport(mode, true);
            } else {
              setAprovStatus('Exportação cancelada pelo usuário.', false);
              showToast('Exportação cancelada.', 'info');
              return;
            }
          }
          
          const msg = errData?.error || 'Erro ao gerar base de aprovação atualizada.';
          setAprovStatus(msg, true);
          showToast(msg, 'danger');
          return;
        }

        const blob = await response.blob();
        if (!blob || !blob.size) {
          const msg = 'Arquivo retornado vazio ou inválido.';
          setAprovStatus(msg, true);
          showToast(msg, 'danger');
          return;
        }

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'base_aprovacao_atualizada.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);

        setAprovStatus('Base de aprovação atualizada gerada com sucesso.', false);
        showToast('Base de aprovação atualizada gerada com sucesso.', 'success');
      } catch (err) {
        const msg = 'Erro ao baixar a base de aprovação atualizada.';
        setAprovStatus(msg, true);
        showToast(msg, 'danger');
      } finally {
        if (aprovRemoveAllBtn) aprovRemoveAllBtn.disabled = false;
        if (aprovRemoveSelectedBtn) aprovRemoveSelectedBtn.disabled = false;
        if (aprovPreviewBtn) aprovPreviewBtn.disabled = false;
      }
    }

    if (aprovRemoveAllBtn) {
      aprovRemoveAllBtn.addEventListener('click', () => {
        callAprovExport('all');
      });
    }
    if (aprovRemoveSelectedBtn) {
      aprovRemoveSelectedBtn.addEventListener('click', () => {
        callAprovExport('selected');
      });
    }
  } catch (e) {
    /* noop */
  }
});
