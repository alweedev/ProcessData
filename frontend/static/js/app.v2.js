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

  async function postFiles(url, filesInput, extra = {}, single = false) {
    const fd = new FormData();
    if (single) {
      if (!filesInput || !filesInput.base || !filesInput.base.files?.[0]) {
        throw new Error("Arquivo base ausente");
      }
      fd.append("base", filesInput.base.files[0]);
      if (filesInput.lista?.files?.[0]) {
        fd.append("lista", filesInput.lista.files[0]);
      }
    } else {
      for (const file of filesInput.files || []) {
        fd.append("files[]", file);
      }
    }
    Object.entries(extra || {}).forEach(([key, value]) =>
      fd.append(key, value)
    );

    const xhr = new XMLHttpRequest();
    const fullUrl =
      window.API_BASE && url.startsWith("/") ? `${window.API_BASE}${url}` : url;
    xhr.open("POST", fullUrl, true);
    xhr.responseType = "blob";

    return new Promise((resolve, reject) => {
      xhr.upload.addEventListener("progress", (event) => {
        if (!event.lengthComputable) return;
        const percent = (event.loaded / event.total) * 100;
        const progressBarId = url.includes("process_cadastro")
          ? "cadastro_progressBar"
          : "inativacao_progressBar";
        const progressBar = document.getElementById(progressBarId);
        if (!progressBar) return;
        progressBar.style.width = `${percent}%`;
        progressBar.setAttribute("aria-valuenow", percent.toFixed(1));
        progressBar.textContent = `${percent.toFixed(1)}%`;
      });

      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300) {
          return resolve({ blob: xhr.response });
        }
        const { status, statusText, response } = xhr;
        if (!response)
          return reject(new Error(`Server ${status}: ${statusText}`));
        const reader = new FileReader();
        reader.onload = () => {
          const text = reader.result || "";
          try {
            const parsed = JSON.parse(text);
            const message =
              parsed.error || parsed.message || text || `Server ${status}`;
            const details = parsed.errors
              ? ` | details: ${JSON.stringify(parsed.errors)}`
              : "";
            reject(new Error(`${message}${details}`));
          } catch (error) {
            reject(new Error(`Server ${status}: ${text || statusText}`));
          }
        };
        reader.onerror = () =>
          reject(new Error(`Server ${status}: ${statusText}`));
        reader.readAsText(response);
      };

      xhr.onerror = () => reject(new Error("Erro de rede"));
      xhr.send(fd);
    });
  }

  async function postFormDataJson(url, formData) {
    const fullUrl =
      window.API_BASE && url.startsWith("/") ? `${window.API_BASE}${url}` : url;
    const response = await fetch(fullUrl, { method: "POST", body: formData });
    const text = await response.text();
    let parsed = {};
    try {
      parsed = text ? JSON.parse(text) : {};
    } catch (error) {
      if (!response.ok) {
        throw new Error(
          `Server ${response.status}: ${text || response.statusText}`
        );
      }
      return {};
    }
    if (!response.ok) {
      const message = parsed.error || parsed.message || response.statusText;
      const details = parsed.errors
        ? ` | details: ${JSON.stringify(parsed.errors)}`
        : "";
      throw new Error(`Server ${response.status}: ${message}${details}`);
    }
    return parsed;
  }

  function showToast(message, type = "success") {
    const toastContainer = document.getElementById("toastContainer");
    if (!toastContainer) return console.warn("toastContainer não encontrado");
    const background =
      type === "success" ? "success" : type === "info" ? "info" : "danger";
    const toast = document.createElement("div");
    toast.className = `toast align-items-center text-white bg-${background} border-0`;
    toast.setAttribute("role", "status");
    toast.setAttribute("aria-live", "polite");
    toast.innerHTML = `<div class="d-flex"><div class="toast-body">${message}</div><button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Fechar"></button></div>`;
    toastContainer.appendChild(toast);
    const instance = new bootstrap.Toast(toast, { delay: 3600 });
    instance.show();
    toast.addEventListener("hidden.bs.toast", () => toast.remove());
  }
  // Expor globalmente para outros scripts (ex: inativacao/index.js)
  try { window.showToast = showToast; } catch(_) {}

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

  // ===== THEME & ACCESSIBILITY CONTROLS =====
  const THEME_KEY = "app_theme";
  const themeToggle = document.getElementById("themeToggle");
  const iconSun = document.getElementById("iconSun");
  const iconMoon = document.getElementById("iconMoon");

  function applyTheme(mode, persist = true) {
    const nextMode = mode === "dark" ? "dark" : "light";
    const isDark = nextMode === "dark";
    document.body.classList.toggle("dark", isDark);
    document.body.setAttribute("data-bs-theme", isDark ? "dark" : "light");
    iconSun?.classList.toggle("d-none", isDark);
    iconMoon?.classList.toggle("d-none", !isDark);
    themeToggle?.setAttribute(
      "aria-label",
      isDark ? "Alternar para tema claro" : "Alternar para tema escuro"
    );
    if (persist) {
      localStorage.setItem(THEME_KEY, nextMode);
    }
    applyRasterInvertToUploadZones();
  }

  if (themeToggle) {
    const storedTheme = localStorage.getItem(THEME_KEY);
    const prefersDark = window.matchMedia
      ? window.matchMedia("(prefers-color-scheme: dark)").matches
      : false;
    const initialTheme = storedTheme || (prefersDark ? "dark" : "light");
    applyTheme(initialTheme, false);
    themeToggle.addEventListener("click", () => {
      const nextMode = document.body.classList.contains("dark")
        ? "light"
        : "dark";
      applyTheme(nextMode, true);
    });
  }

  const MOTION_KEY = "motion_pref";
  const motionToggle = document.getElementById("motionToggle");
  const motionLabel = motionToggle?.querySelector("span");

  function applyMotionPreference(reduced, persist = true) {
    document.body.classList.toggle("reduce-motion", !!reduced);
    if (motionLabel) {
      motionLabel.textContent = reduced
        ? "Animações: Reduzidas"
        : "Animações: Ativas";
    }
    if (persist) {
      localStorage.setItem(MOTION_KEY, reduced ? "reduce" : "full");
    }
  }

  if (motionToggle) {
    const storedMotion = localStorage.getItem(MOTION_KEY);
    const prefersReduced = window.matchMedia
      ? window.matchMedia("(prefers-reduced-motion: reduce)").matches
      : false;
    const initialMotion = storedMotion
      ? storedMotion === "reduce"
      : prefersReduced;
    applyMotionPreference(initialMotion, false);
    motionToggle.addEventListener("click", () => {
      const nextValue = !document.body.classList.contains("reduce-motion");
      applyMotionPreference(nextValue, true);
    });
  }

  // ===== API STATUS MONITOR =====
  const apiStatusBtn = document.getElementById("apiStatusBtn");
  const apiStatusIcon = document.getElementById("apiStatusIcon");
  const apiStatusText = document.getElementById("apiStatusText");
  let apiStatusInterval = null;
  let lastApiState = "unknown";

  const apiIcons = {
    // Stylized hub (central circle) with connected nodes for online
    online: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="3"/><path stroke-linecap="round" stroke-linejoin="round" d="M12 9V5M12 19v-4M9 12H5M19 12h-4M9.6 9.6l-2.8-2.8M16.4 9.6l2.8-2.8M9.6 14.4l-2.8 2.8M16.4 14.4l2.8 2.8"/>',
    // Offline: broken link chain with warning slash
    offline: '<circle cx="12" cy="12" r="9"/><path stroke-linecap="round" stroke-linejoin="round" d="M8.5 13.5l3-3m1 1.5l2.2 2.2M14.5 10.5L16 9m-8 6l-1.5 1.5M9 8.8L7.5 7.3"/><path stroke-linecap="round" stroke-linejoin="round" d="M6 6l12 12"/>',
    // Checking: spinner arc + pulse dot
    checking: '<circle cx="12" cy="12" r="9"/><path stroke-linecap="round" stroke-linejoin="round" d="M12 3a9 9 0 019 9"/><circle cx="12" cy="12" r="2"/>'
  };

  function setApiStatusVisual(state) {
    if (!apiStatusBtn || !apiStatusIcon || !apiStatusText) return;
    apiStatusBtn.classList.remove("api-online", "api-offline", "api-checking");
    apiStatusBtn.classList.add(`api-${state}`);
    apiStatusIcon.innerHTML = apiIcons[state] || apiIcons.offline;
    apiStatusIcon.classList.toggle("spin-rotating", state === "checking");
    const textMap = {
      online: "API Online",
      offline: "API Offline",
      checking: "Verificando...",
    };
    apiStatusText.textContent = textMap[state] || "API";
    apiStatusBtn.removeAttribute("disabled");
  }

  async function pingApi(manual = false) {
    if (!apiStatusBtn || !apiStatusIcon) return;
    setApiStatusVisual("checking");

    const base = window.API_BASE ? window.API_BASE.replace(/\/$/, "") : "";

    const candidates = [
      `${base}/health`,
      `${base}/api/health`,
      base ? `${base}/` : "/",
    ];

    async function tryFetch(url) {
      const controller = window.AbortController ? new AbortController() : null;
      const timeoutId = controller ? setTimeout(() => controller.abort(), 6000) : null;
      try {
        const res = await fetch(url, { method: "GET", signal: controller?.signal });
        if (timeoutId) clearTimeout(timeoutId);
        return res.ok;
      } catch (e) {
        if (timeoutId) clearTimeout(timeoutId);
        return false;
      }
    }

    for (const url of candidates) {
      // Skip duplicate URLs
      if (!url || (typeof url === "string" && url.endsWith("//"))) continue;
      // If one succeeds, mark online and stop
      /* eslint-disable no-await-in-loop */
      const ok = await tryFetch(url);
      if (ok) {
        handleApiResult(true, manual);
        return;
      }
    }
    handleApiResult(false, manual);
  }

  function handleApiResult(isOnline, notify) {
    const state = isOnline ? "online" : "offline";
    setApiStatusVisual(state);
    if (notify || lastApiState !== state) {
      showToast(
        isOnline ? "API Online" : "API Offline",
        isOnline ? "success" : "danger"
      );
    }
    lastApiState = state;
  }

  if (apiStatusBtn && apiStatusIcon && apiStatusText) {
    apiStatusBtn.addEventListener("click", () => pingApi(true));
    pingApi(false);
    apiStatusInterval = setInterval(() => pingApi(false), 45000);
    window.addEventListener("beforeunload", () => {
      if (apiStatusInterval) clearInterval(apiStatusInterval);
    });
  }

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
        addToHistory(
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

  /**
   * Gera HTML amigável para exibir correspondências que já estão INATIVAS
   * Espera um objeto onde cada chave é um tipo (ex: 'cpf','exact','token')
   * e o valor é um array de objetos (registros). Limita a 10 linhas por grupo.
   */
  function renderInactiveMatchesHtml(inactiveObj) {
    if (!inactiveObj || typeof inactiveObj !== "object") return "";
    const parts = [];
    parts.push(
      `<div class="inativos-controls"><button type="button" class="btn btn-sm btn-outline-primary" id="inativos_expand_all">Expandir todos</button><button type="button" class="btn btn-sm btn-outline-secondary d-none" id="inativos_collapse_all">Colapsar todos</button><div class="ms-auto small text-muted">Clique em cada grupo para ver a amostra</div></div>`
    );
    parts.push(
      '<div class="accordion inativos-accordion" id="inativosAccordion">'
    );
    let idx = 0;
    for (const [group, arr] of Object.entries(inactiveObj)) {
      if (!Array.isArray(arr) || arr.length === 0) {
        idx++;
        continue;
      }
      const sample = arr.slice(0, 10);
      const preferred = [
        "CPF",
        "CPFdigits",
        "NomeCompleto",
        "Nome Normalizado",
        "Status",
        "Email",
        "Departamento",
        "NomeEmpresa",
        "UserId",
      ];
      const keySet = new Set();
      sample.forEach((row) =>
        Object.keys(row || {}).forEach((key) => keySet.add(key))
      );
      const cols = [];
      preferred.forEach((key) => {
        if (keySet.has(key)) {
          cols.push(key);
          keySet.delete(key);
        }
      });
      Array.from(keySet)
        .slice(0, 8)
        .forEach((key) => cols.push(key));

      const headingId = `inativosHeading${idx}`;
      const collapseId = `inativosCollapse${idx}`;
      let html = "";
      html += '<div class="accordion-item">';
      html += `<h2 class="accordion-header" id="${headingId}">`;
      html += `<button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}" aria-expanded="false" aria-controls="${collapseId}">Inativos - ${escapeHtml(
        group
      )} <span class="ms-2 small text-muted">(${
        arr.length
      } total)</span></button>`;
      html += "</h2>";
      html += `<div id="${collapseId}" class="accordion-collapse collapse" aria-labelledby="${headingId}" data-bs-parent="#inativosAccordion">`;
      html += '<div class="accordion-body">';
      html += '<div class="inativos-table-wrapper">';
      html +=
        '<table class="table table-sm table-bordered"><thead class="table-light"><tr>' +
        cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("") +
        "</tr></thead><tbody>";
      for (const row of sample) {
        html +=
          "<tr>" +
          cols.map((c) => `<td>${escapeHtml(row?.[c] ?? "")}</td>`).join("") +
          "</tr>";
      }
      html += "</tbody></table>";
      if (arr.length > sample.length) {
        html += `<div class="small text-muted">Mostrando ${sample.length} de ${arr.length} itens. Exporte para análise completa.</div>`;
      }
      html += "</div>";
      html += "</div></div>";
      html += "</div>";
      parts.push(html);
      idx++;
    }
    parts.push("</div>");
    parts.push(
      `<script>document.addEventListener('DOMContentLoaded',function(){const expandBtn=document.getElementById('inativos_expand_all');const collapseBtn=document.getElementById('inativos_collapse_all');if(expandBtn){expandBtn.addEventListener('click',()=>{document.querySelectorAll('#inativosAccordion .accordion-collapse').forEach(c=>{if(!c.classList.contains('show')){const bs=bootstrap.Collapse.getOrCreateInstance(c,{toggle:false});bs.show();}});expandBtn.classList.add('d-none');collapseBtn.classList.remove('d-none');});}if(collapseBtn){collapseBtn.addEventListener('click',()=>{document.querySelectorAll('#inativosAccordion .accordion-collapse.show').forEach(c=>{const bs=bootstrap.Collapse.getOrCreateInstance(c,{toggle:false});bs.hide();});collapseBtn.classList.add('d-none');expandBtn.classList.remove('d-none');});}});</script>`
    );
    return parts.join("");
  }

  // Upload de lista de inativação removido – sem drag & drop de lista
  // Inativação - Button click
  const inativacaoBtn = document.getElementById("inativacao_btn");
  // flag para evitar envios duplicados
  let inativacaoInProgress = false;
  if (inativacaoBtn) {
    inativacaoBtn.addEventListener("click", async () => {
      // Short-circuit legacy modal flow if enhanced inativação UX is active
      if (window.INATIVACAO_ENHANCED) {
        // Enhanced script handles buscar + geração; prevent duplicate modal
        return;
      }
      if (inativacaoInProgress) return;
      inativacaoInProgress = true;
      inativacaoBtn.classList.add("btn-processing");
      const base = document.getElementById("inativacao_base");
      const lista = document.getElementById("inativacao_lista");
      const listaText = document.getElementById("lista_text"); // Novo campo
      const status = document.getElementById("inativacao_status");
      const debug = document.getElementById("inativacao_debug");
      const loading = document.getElementById("inativacao_loading");
      const progress = document.getElementById("inativacao_progress");

      if (!status || !debug || !loading || !progress)
        console.error("Elementos de status inativação não encontrados!");
      setStatus(
        status,
        '<span class="spinner-border spinner-border-sm text-primary me-2" role="status"></span>Processando...'
      );
      debug.textContent = "";
      loading.classList.remove("d-none");
      progress.classList.remove("d-none");

      try {
        if (!base.files[0]) throw new Error("Envie a base");
        const extraData = {};
        if (listaText && listaText.value) {
          extraData.lista_text = listaText.value;
        } else if (!lista.files[0]) {
          throw new Error("Envie a lista ou insira os nomes/CPFs");
        }
        // Preview: pedir ao servidor quantas correspondências serão geradas
        const fdPreview = new FormData();
        fdPreview.append("base", base.files[0]);
        if (lista.files[0]) fdPreview.append("lista", lista.files[0]);
        if (extraData.lista_text)
          fdPreview.append("lista_text", extraData.lista_text);
        // anexar parâmetros de fuzzy
        const useFuzzyCheckbox = document.getElementById("use_fuzzy");
        const fuzzyCutoffInput = document.getElementById("fuzzy_cutoff");
        if (useFuzzyCheckbox)
          fdPreview.append(
            "use_fuzzy",
            useFuzzyCheckbox.checked ? "true" : "false"
          );
        if (fuzzyCutoffInput)
          fdPreview.append("fuzzy_cutoff", fuzzyCutoffInput.value || "0.90");
        let preview;
        try {
          preview = await postFormDataJson(
            "/api/preview_inativacao",
            fdPreview
          );
        } catch (errPreview) {
          // Extrair mensagem detalhada do erro do servidor
          let errorMsg = "Erro ao gerar preview";
          if (errPreview && errPreview.message) {
            errorMsg = errPreview.message;
            // Se a mensagem já contém "Server 400:" ou "Server 500:", não duplicar
            if (!errorMsg.includes("Server")) {
              errorMsg = `Erro no preview: ${errorMsg}`;
            }
          }
          throw new Error(errorMsg);
        }
        try {
          console.debug("preview payload", preview);
        } catch (e) {}

        // Mostrar confirmação com contagem e amostra (se houver)
        // Fallback: se o backend não retornou out_df/sample mas trouxe
        // stats.inactive_matches, gerar sample a partir desse objeto para
        // manter a experiência anterior (mostrar tabela para confirmação).
        let count = preview.count || 0;
        let sample = preview.sample || [];

        if (
          (!sample || sample.length === 0) &&
          preview.stats &&
          preview.stats.inactive_matches
        ) {
          try {
            const inactiveObj = preview.stats.inactive_matches || {};
            const all = [];
            for (const [type, arr] of Object.entries(inactiveObj)) {
              if (Array.isArray(arr)) {
                for (const it of arr) {
                  all.push(Object.assign({ match_type: type }, it));
                }
              }
            }
            if (all.length) {
              count = all.length;
              sample = all.slice(0, 10);
              // persistir para consistência com código que ainda usa preview.sample
              preview.sample = sample;
            }
          } catch (e) {
            console.debug(
              "fallback build sample from inactive_matches failed",
              e
            );
          }
        }

        // Preferir stats.total_matches quando presente (fonte de verdade do backend)
        const displayCount =
          preview.stats && typeof preview.stats.total_matches === "number"
            ? preview.stats.total_matches
            : count;

        let infoHtml = `<p>Foram encontradas <strong>${escapeHtml(
          displayCount
        )}</strong> correspondência(s).</p>`;
        if (sample && sample.length) {
          const keys = Object.keys(sample[0]).slice(0, 6);
          const rowsHtml = sample
            .slice(0, 10)
            .map((row) => {
              return `<tr>${keys
                .map((k) => `<td>${escapeHtml(row[k] || "")}</td>`)
                .join("")}</tr>`;
            })
            .join("");
          infoHtml += `
            <div style="max-height:200px;overflow:auto;margin-top:8px;">
              <table class="table table-sm table-bordered"><thead><tr>
                ${keys.map((k) => `<th>${escapeHtml(k)}</th>`).join("")}
              </tr></thead><tbody>
                ${rowsHtml}
              </tbody></table>
            </div>`;
        }

        // Nova seção: Tabela detalhada para conferência antes do download
        if (Array.isArray(preview.records) && preview.records.length) {
          // escolher colunas prioritárias e limitar o resto
          const preferred = [
            "Login",
            "UserId",
            "NomeCompleto",
            "Email",
            "Departamento",
            "NomeEmpresa",
            "CodigoCCustoEmpresa",
            "Solicitante",
          ];
          const cols = Array.isArray(preview.columns)
            ? preview.columns.slice()
            : Object.keys(preview.records[0] || {});
          const colsToShow = [];
          preferred.forEach((c) => {
            if (cols.includes(c)) {
              colsToShow.push(c);
              cols.splice(cols.indexOf(c), 1);
            }
          });
          // completar com primeiras colunas restantes (até 8 totais)
          while (colsToShow.length < 8 && cols.length)
            colsToShow.push(cols.shift());
          const initialRows = preview.records.slice(0, 200);
          const headerHtml = colsToShow
            .map((c) => `<th scope="col">${escapeHtml(c)}</th>`)
            .join("");
          const bodyHtml = initialRows
            .map(
              (row) =>
                `<tr>${colsToShow
                  .map((c) => `<td>${escapeHtml(row[c] ?? "")}</td>`)
                  .join("")}</tr>`
            )
            .join("");
          infoHtml += `
            <div class="mt-3">
              <div class="d-flex align-items-center mb-2 gap-2">
                <input id="preview_search" class="form-control form-control-sm" placeholder="Filtrar na prévia..." />
                <button id="export_preview_csv" class="btn btn-sm btn-outline-success">Exportar Prévia CSV</button>
              </div>
              <div style="max-height:50vh; overflow:auto;">
                <table class="table table-sm table-hover table-bordered">
                  <thead class="table-light"><tr>${headerHtml}</tr></thead>
                  <tbody id="preview_table_body">${bodyHtml}</tbody>
                </table>
              </div>
              <div class="small text-muted">Mostrando ${initialRows.length} de ${preview.records.length} linhas.</div>
            </div>`;
        }
        // attach stats breakdown (if returned by backend)
        if (preview.stats) {
          // se houver correspondências que já estão INATIVAS, exibir uma tabela amigável
          if (preview.stats.inactive_matches) {
            const inact = preview.stats.inactive_matches;
            // bloco de aviso resumido com contagens
            try {
              const counts = Object.fromEntries(
                Object.entries(inact).map(([k, v]) => [
                  k,
                  Array.isArray(v) ? v.length : 0,
                ])
              );
              const summary = `<div class="alert alert-warning"><strong>Atenção:</strong> Foram encontradas correspondências que já estão <strong>INATIVAS</strong> na base.<div class="mt-2 small">Resumo: ${Object.entries(
                counts
              )
                .map(([k, c]) => `${k}: ${c}`)
                .join(" | ")}</div></div>`;
              infoHtml = summary + infoHtml;
            } catch (e) {
              infoHtml =
                `<div class="alert alert-warning"><strong>Atenção:</strong> Foram encontradas correspondências já INATIVAS.</div>` +
                infoHtml;
            }
            // agora adicionar tabelas legíveis para cada grupo
            infoHtml += renderInactiveMatchesHtml(
              preview.stats.inactive_matches
            );
          }
          // e por fim as estatísticas resumidas
          infoHtml += renderPreviewStatsHtml(preview.stats);
        }

        // adicionar botão de exportar CSV acima do conteúdo para auditoria offline
        const exportControls = `<div class="d-flex mb-2"><button id="export_inactives_csv" class="btn btn-sm btn-success me-2">Exportar CSV</button><div class="ms-auto small text-muted">Exporta todas as correspondências inativas</div></div>`;
        infoHtml = exportControls + infoHtml;

        // Helper: converte array de objetos em CSV (keys unificadas)
        function objectsToCsv(objs) {
          if (!objs || !objs.length) return "";
          const keys = Array.from(
            new Set(objs.flatMap((o) => Object.keys(o || {})))
          );
          const escape = (v) =>
            '"' +
            String(v === null || v === undefined ? "" : v).replace(/"/g, '""') +
            '"';
          const header = keys.map((k) => escape(k)).join(",");
          const lines = objs.map((o) =>
            keys.map((k) => escape(o[k] ?? "")).join(",")
          );
          return [header].concat(lines).join("\r\n");
        }

        // Helper: SweetAlert fallback se CDN falhar
        async function confirmWithFallback(options) {
          try {
            if (window.Swal && typeof Swal.fire === "function") {
              return await Swal.fire(options);
            }
          } catch (e) {
            /* fall through to confirm() */
          }
          const txt =
            options && (options.title || options.html)
              ? `${options.title || ""}\n\n${
                  options.html
                    ? typeof options.html === "string"
                      ? options.html.replace(/<[^>]*>/g, "")
                      : ""
                    : ""
                }`.trim()
              : "Confirmar?";
          const ok = window.confirm(txt);
          return { isConfirmed: !!ok };
        }

        const confirmed = await confirmWithFallback({
          title: "Confirmar inativação",
          html: infoHtml,
          showCancelButton: true,
          confirmButtonText: "Gerar inativação",
          cancelButtonText: "Cancelar",
          width: "95vw",
          customClass: { popup: "swal2-large-modal" },
          didOpen: (popup) => {
            try {
              const btn = popup.querySelector("#export_inactives_csv");
              if (btn) {
                btn.addEventListener("click", () => {
                  try {
                    const inactive =
                      preview.stats && preview.stats.inactive_matches
                        ? preview.stats.inactive_matches
                        : {};
                    // concatenar todas as listas em uma só com uma coluna extra 'match_type'
                    const all = [];
                    for (const [type, arr] of Object.entries(inactive || {})) {
                      if (Array.isArray(arr)) {
                        for (const it of arr) {
                          all.push(Object.assign({ match_type: type }, it));
                        }
                      }
                    }
                    if (!all.length) {
                      showToast(
                        "Não há correspondências inativas para exportar.",
                        "info"
                      );
                      return;
                    }
                    const csv = objectsToCsv(all);
                    const blob = new Blob([csv], {
                      type: "text/csv;charset=utf-8;",
                    });
                    downloadBlob(blob, "inactive_matches.csv");
                    showToast("CSV de correspondências gerado.", "success");
                  } catch (e) {
                    showToast(
                      "Falha ao gerar CSV: " + (e.message || e),
                      "danger"
                    );
                  }
                });
              }
              // Wiring da tabela de prévia (pesquisa e export)
              try {
                const searchEl = popup.querySelector("#preview_search");
                const tbody = popup.querySelector("#preview_table_body");
                const exportPreviewBtn = popup.querySelector(
                  "#export_preview_csv"
                );
                if (
                  tbody &&
                  Array.isArray(preview.records) &&
                  preview.records.length
                ) {
                  // Reconstituir as colunas exibidas a partir do thead
                  const ths = popup.querySelectorAll("thead tr th");
                  const shownCols = Array.from(ths).map((th) => th.textContent);
                  const rowsAll = preview.records.slice();

                  const render = (rows) => {
                    const limited = rows.slice(0, 500);
                    tbody.innerHTML = limited
                      .map(
                        (row) =>
                          `<tr>${shownCols
                            .map((c) => `<td>${row[c] ?? ""}</td>`)
                            .join("")}</tr>`
                      )
                      .join("");
                  };
                  if (searchEl) {
                    searchEl.addEventListener("input", (e) => {
                      const term = (e.target.value || "").toLowerCase();
                      if (!term) {
                        render(rowsAll);
                        return;
                      }
                      const filtered = rowsAll.filter((r) =>
                        shownCols.some((c) =>
                          String(r[c] ?? "")
                            .toLowerCase()
                            .includes(term)
                        )
                      );
                      render(filtered);
                    });
                  }
                  render(rowsAll);

                  if (exportPreviewBtn) {
                    exportPreviewBtn.addEventListener("click", () => {
                      const csv = objectsToCsv(preview.records);
                      const blob = new Blob([csv], {
                        type: "text/csv;charset=utf-8;",
                      });
                      downloadBlob(blob, "preview_inativacao.csv");
                      showToast("CSV da prévia exportado.", "success");
                    });
                  }
                }
              } catch (e) {
                console.debug("preview table wiring error", e);
              }
            } catch (e) {
              console.debug("didOpen export control error", e);
            }
          },
        });
        if (!confirmed.isConfirmed) {
          setStatus(status, "");
          loading.classList.add("d-none");
          progress.classList.add("d-none");
          inativacaoBtn.classList.remove("btn-processing");
          return; // usuário cancelou
        }

        // Se confirmado, seguir com a geração do arquivo
        // adicionar os parâmetros de fuzzy ao extraData também
        if (useFuzzyCheckbox)
          extraData.use_fuzzy = useFuzzyCheckbox.checked ? "true" : "false";
        if (fuzzyCutoffInput)
          extraData.fuzzy_cutoff = fuzzyCutoffInput.value || "0.90";

        const resp = await postFiles(
          "/api/process_inativacao",
          { base, lista: lista.files[0] ? lista : null },
          extraData,
          true
        );
        if (!resp.blob || resp.blob.size === 0)
          throw new Error("Arquivo gerado inválido");
        downloadBlob(resp.blob, "saida_inativacao.xlsx");
        setStatus(status, '<span class="text-success">✔ Concluído</span>');
        showToast("Inativação processada!", "success");
        addToHistory(
          `Inativação gerada: Base ${base.files[0].name}, Lista ${
            lista.files[0]?.name || "texto"
          } - ${new Date().toLocaleString()}`
        );
      } catch (err) {
        setStatus(status, "");
        // Melhorar apresentação de erros do backend
        let errorText = "Erro desconhecido";
        if (err && err.message) {
          errorText = err.message;
          // Se a mensagem contém JSON, tentar extrair campo "error"
          try {
            const match = err.message.match(/\{.*\}/);
            if (match) {
              const obj = JSON.parse(match[0]);
              if (obj.error) errorText = obj.error;
            }
          } catch (e) {
            /* JSON parse failed, use raw message */
          }
        }
        debug.textContent = "Erro: " + errorText;
        showToast(`Erro ao processar inativação: ${errorText}`, "danger");
      }
      loading.classList.add("d-none");
      progress.classList.add("d-none");
      inativacaoBtn.classList.remove("btn-processing");
      inativacaoInProgress = false;
      document.getElementById("inativacao_progressBar").style.width = "0%";
      document.getElementById("inativacao_progressBar").textContent = "";
    });
  }

  // Atalho: Ctrl+Enter no textarea #lista_text dispara o envio de Inativação
  try {
    const listaTextEl = document.getElementById("lista_text");
    if (listaTextEl) {
      listaTextEl.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
          // evita duplo envio; dispara clique no botão
          e.preventDefault();
          inativacaoBtn && inativacaoBtn.click();
        }
      });
    }
  } catch (e) {
    /* ignore */
  }

  // persist fuzzy preferences
  try {
    const useFuzzy = document.getElementById("use_fuzzy");
    const fuzzyCutoff = document.getElementById("fuzzy_cutoff");
    if (useFuzzy && fuzzyCutoff) {
      const stored = JSON.parse(localStorage.getItem("fuzzy_prefs") || "{}");
      if (stored.use_fuzzy !== undefined) useFuzzy.checked = stored.use_fuzzy;
      if (stored.fuzzy_cutoff) fuzzyCutoff.value = stored.fuzzy_cutoff;
      useFuzzy.addEventListener("change", () => {
        localStorage.setItem(
          "fuzzy_prefs",
          JSON.stringify({
            use_fuzzy: useFuzzy.checked,
            fuzzy_cutoff: fuzzyCutoff.value,
          })
        );
      });
      fuzzyCutoff.addEventListener("change", () => {
        localStorage.setItem(
          "fuzzy_prefs",
          JSON.stringify({
            use_fuzzy: useFuzzy.checked,
            fuzzy_cutoff: fuzzyCutoff.value,
          })
        );
      });
    }
  } catch (e) {
    /* ignore */
  }

  // Persistir escolhas de cadastro (login choice e fluxo)
  try {
    const loginChoiceEl = document.getElementById("cadastro_login_choice");
    const fluxoEl = document.getElementById("cadastro_fluxo");
    const storedPrefs = JSON.parse(
      localStorage.getItem("cadastro_prefs") || "{}"
    );
    if (loginChoiceEl && storedPrefs.login_choice)
      loginChoiceEl.value = storedPrefs.login_choice;
    if (fluxoEl && storedPrefs.fluxo) fluxoEl.value = storedPrefs.fluxo;
    if (loginChoiceEl)
      loginChoiceEl.addEventListener("change", () => {
        const prefs = JSON.parse(
          localStorage.getItem("cadastro_prefs") || "{}"
        );
        prefs.login_choice = loginChoiceEl.value;
        localStorage.setItem("cadastro_prefs", JSON.stringify(prefs));
      });
    if (fluxoEl)
      fluxoEl.addEventListener("change", () => {
        const prefs = JSON.parse(
          localStorage.getItem("cadastro_prefs") || "{}"
        );
        prefs.fluxo = fluxoEl.value;
        localStorage.setItem("cadastro_prefs", JSON.stringify(prefs));
      });
  } catch (e) {
    /* ignore */
  }

  // Histórico (modernizado)
  const historicoTbody = document.getElementById("historico_tbody");
  const historicoSearch = document.getElementById("historico_search");
  const historicoSummary = document.getElementById("historico_summary");
  const exportCsvBtn = document.getElementById("historico_export_csv");
  const exportJsonBtn = document.getElementById("historico_export_json");
  const clearHistoryBtn = document.getElementById("clearHistoryBtn");
  const HISTORY_KEY = "history_v2"; // novo formato

  function migrateOldHistory() {
    // Se existir "history" simples (array de strings) e não existir history_v2
    const oldRaw = localStorage.getItem("history_v2");
    if (oldRaw) return; // já migrado
    const legacy = JSON.parse(localStorage.getItem("history") || "[]");
    if (Array.isArray(legacy) && legacy.length) {
      const now = Date.now();
      const migrated = legacy.map((text, i) => ({
        ts: now - (legacy.length - i) * 1000,
        text: String(text),
      }));
      localStorage.setItem(HISTORY_KEY, JSON.stringify(migrated.slice(-200)));
      // opcionalmente remover legacy
      try { localStorage.removeItem("history"); } catch(_) {}
    }
  }
  migrateOldHistory();

  function getHistory() {
    const arr = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
    return Array.isArray(arr) ? arr : [];
  }
  function setHistory(arr) {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(arr.slice(-500))); // limite maior
  }
  function addToHistory(action) {
    const history = getHistory();
    history.push({ ts: Date.now(), text: String(action) });
    setHistory(history);
    renderHistory();
  }
  function formatTs(ts) {
    try {
      const d = new Date(ts);
      const pad = (n) => String(n).padStart(2, "0");
      return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
    } catch (_) {
      return "—";
    }
  }
  let filteredHistory = [];
  function renderHistory() {
    const all = getHistory();
    const term = (historicoSearch?.value || "").trim().toLowerCase();
    filteredHistory = term
      ? all.filter((h) => h.text.toLowerCase().includes(term))
      : all.slice();
    if (historicoTbody) {
      historicoTbody.innerHTML = filteredHistory
        .slice()
        .reverse() // mostrar recentes primeiro
        .map(
          (h) => `<tr>
            <td style="white-space:nowrap;">${formatTs(h.ts)}</td>
            <td>${escapeHtml(h.text)}</td>
          </tr>`
        )
        .join("");
    }
    if (historicoSummary) {
      historicoSummary.textContent = `Itens: ${filteredHistory.length} (total armazenado: ${all.length})`;
    }
  }
  historicoSearch?.addEventListener("input", () => renderHistory());
  clearHistoryBtn?.addEventListener("click", () => {
    setHistory([]);
    renderHistory();
    showToast("Histórico limpo!", "success");
  });
  exportCsvBtn?.addEventListener("click", () => {
    if (!filteredHistory.length) {
      showToast("Nada para exportar.", "warning");
      return;
    }
    const header = "data_hora,acao";
    const rows = filteredHistory
      .slice()
      .reverse()
      .map((h) => `${formatTs(h.ts).replace(/,/g, " ")},"${h.text.replace(/"/g, '""')}"`);
    const csv = [header].concat(rows).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `historico_${Date.now()}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
    showToast("CSV exportado.", "success");
  });
  exportJsonBtn?.addEventListener("click", () => {
    if (!filteredHistory.length) {
      showToast("Nada para exportar.", "warning");
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
    showToast("JSON exportado.", "success");
  });
  renderHistory();
  function wireAnaliseHelp() {
    const btn = document.getElementById("analise_help_btn");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const html = `
        <div class="text-start">
          <p class="mb-2 fw-semibold">Guia rápido do workspace</p>
          <ol class="ps-3 small">
            <li><strong>Envie</strong> a planilha arrastando ou clicando em <em>Selecionar arquivo</em>.</li>
            <li>O processamento é incremental: você pode navegar enquanto carrega.</li>
            <li>A <strong>busca</strong> filtra qualquer coluna instantaneamente.</li>
            <li>Ajuste <strong>quebra de texto</strong>, <strong>densidade</strong> e <strong>fixe</strong> o resumo conforme sua preferência.</li>
            <li>O <strong>CSV exportado</strong> respeita filtros ativos e aplica limites para manter a performance.</li>
          </ol>
          <p class="small text-muted mb-0">Tudo acontece no navegador – nenhum dado sensível sai da sua máquina.</p>
        </div>`;
      if (window.Swal?.fire) {
        Swal.fire({ title: "Workspace Analítico", html, confirmButtonText: "Continuar", width: 540 });
      } else {
        alert("Envie o Excel, ajuste preferências e exporte quando quiser.");
      }
    });
  }

  function wireCadastroHelp() {
    const btn = document.getElementById("cadastro_help_btn");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const html = `
        <div class="text-start">
          <p class="mb-2 fw-semibold">Como usar o cadastro em lote</p>
          <ol class="ps-3 small">
            <li>Carregue a planilha Excel (<strong>.xlsx/.xls</strong>) já preenchida.</li>
            <li>Escolha o <strong>tipo de login</strong> (CPF ou E-mail) e o <strong>fluxo</strong> (SELF ou FRONT).</li>
            <li>Clique em <em>Gerar</em> para processar e obter o arquivo pronto para carga.</li>
            <li>Acompanhe e recupere execuções no <strong>Histórico</strong> quando precisar.</li>
          </ol>
          <p class="small text-muted mb-0">Dica: revise os dados antes de enviar para evitar retrabalho.</p>
        </div>`;
      if (window.Swal?.fire) {
        Swal.fire({ title: "Carga cadastro", html, confirmButtonText: "Fechar", width: 500 });
      } else {
        alert("Envie a base, escolha login e fluxo, e gere o arquivo.");
      }
    });
  }

  function wireInativacaoHelp() {
    const btn = document.getElementById("inativacao_help_btn");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const html = `
        <div class="text-start">
          <p class="mb-2 fw-semibold">Passos para inativação rápida</p>
          <ol class="ps-3 small">
            <li><strong>Envie</strong> a base de usuários em Excel (.xlsx).</li>
            <li>Cole <strong>CPFs</strong> (com ou sem máscara), <strong>nomes completos</strong> ou <strong>e-mails</strong> – um por linha.</li>
            <li>Use <em>Buscar</em> para conferir quem será afetado e estados atuais.</li>
            <li>Finalize em <strong>Gerar</strong> para baixar o relatório da inativação.</li>
          </ol>
          <p class="small text-muted mb-0">Atalho: Ctrl+Enter no campo dispara a busca imediatamente. Suporte completo a e-mail para localizar usuários pelo endereço eletrônico.</p>
        </div>`;
      if (window.Swal?.fire) {
        Swal.fire({ title: "Inativação", html, confirmButtonText: "Fechar", width: 520 });
      } else {
        alert("Envie base, cole CPFs/nomes/e-mails, busque e gere o arquivo.");
      }
    });
  }

  // Wired help buttons
  wireAnaliseHelp();
  wireCadastroHelp();
  wireInativacaoHelp();

  // Cadastro: botão Limpar seleção de arquivos
  try {
    const clearCadastroBtn = document.getElementById('cadastro_clear_btn');
    const cadastroFiles = document.getElementById('cadastro_files');
    const cadastroFeedback = document.getElementById('cadastro_uploadFeedback');
    if (cadastroFiles && cadastroFeedback) {
      cadastroFiles.addEventListener('change', () => {
        if (cadastroFiles.files && cadastroFiles.files.length) {
          const names = Array.from(cadastroFiles.files).map(f => f.name).join(', ');
          setBannerContent(cadastroFeedback, `<span class="upload-check" aria-hidden="true"><i class="bi bi-check-lg text-success" style="font-size:1.15rem; line-height:1; display:inline-block;"></i></span>${names}`);
        } else {
          clearBanner(cadastroFeedback);
        }
      });
    }
    if (clearCadastroBtn) {
      clearCadastroBtn.addEventListener('click', () => {
        const files = document.getElementById('cadastro_files');
        const names = document.getElementById('cadastro_fileNames');
        const status = document.getElementById('cadastro_status');
        const debug = document.getElementById('cadastro_debug');
        const loading = document.getElementById('cadastro_loading');
        const progress = document.getElementById('cadastro_progress');
        const progressBar = document.getElementById('cadastro_progressBar');
        try { if (files) files.value = ''; } catch(_) {}
        try { if (names) names.innerHTML = ''; } catch(_) {}
        try { if (status) status.innerHTML = ''; } catch(_) {}
        try { if (debug) debug.textContent = ''; } catch(_) {}
        try { if (loading) loading.classList.add('d-none'); } catch(_) {}
        try { if (progress) progress.classList.add('d-none'); } catch(_) {}
        try { if (progressBar) { progressBar.style.width = '0%'; progressBar.textContent = ''; progressBar.setAttribute('aria-valuenow','0'); } } catch(_) {}
        try { if (cadastroFeedback) clearBanner(cadastroFeedback); } catch(_) {}
        showToast('Seleção de cadastro limpa.', 'info');
      });
    }

    // Cadastro: processar planilha (ativar botão "Gerar")
    const cadastroBtn = document.getElementById('cadastro_btn');
    let cadastroInProgress = false;
    if (cadastroBtn) {
      cadastroBtn.addEventListener('click', async () => {
        if (cadastroInProgress) return;
        cadastroInProgress = true;
        cadastroBtn.classList.add('btn-processing');

        const files = document.getElementById('cadastro_files');
        const status = document.getElementById('cadastro_status');
        const debug = document.getElementById('cadastro_debug');
        const loading = document.getElementById('cadastro_loading');
        const progress = document.getElementById('cadastro_progress');
        const loginChoiceEl = document.getElementById('cadastro_login_choice');
        const fluxoEl = document.getElementById('cadastro_fluxo');

        try {
          if (!status || !debug || !loading || !progress) {
            console.error('Elementos de status do cadastro não encontrados');
          }
          setStatus(status, '<span class="spinner-border spinner-border-sm text-primary me-2" role="status"></span>Processando...');
          debug.textContent = '';
          loading.classList.remove('d-none');
          progress.classList.remove('d-none');

          if (!files || !files.files || !files.files.length) {
            throw new Error('Selecione pelo menos um arquivo');
          }

          const loginChoice = loginChoiceEl ? loginChoiceEl.value : 'CPF';
          const fluxo = fluxoEl ? fluxoEl.value : 'SELF';
          const extraData = { login_choice: loginChoice, fluxo };
          if (fluxo === 'SELF') {
            extraData.vip = 'N';
            extraData.viajanteMasterNacional = 'N';
            extraData.viajanteMasterInternacional = 'N';
            extraData.solicitanteMaster = 'N';
            extraData.masterAdiantamento = 'N';
            extraData.masterReembolso = 'N';
          } else if (fluxo === 'FRONT') {
            extraData.vip = 'N';
            extraData.viajanteMasterNacional = 'S';
            extraData.viajanteMasterInternacional = 'S';
            extraData.solicitanteMaster = 'N';
            extraData.masterAdiantamento = 'N';
            extraData.masterReembolso = 'N';
          }

          const resp = await postFiles('/api/process_cadastro', files, extraData);
          if (!resp.blob || resp.blob.size === 0) throw new Error('Arquivo gerado inválido');
          downloadBlob(resp.blob, 'saida_cadastro.xlsx');
          setStatus(status, '<span class="text-success">✔ Concluído</span>');
          showToast(`Cadastro processado! Opções: ${loginChoice}, ${fluxo}.`, 'success');
          try { addToHistory(`Cadastro gerado: ${files.files[0].name} - ${new Date().toLocaleString('pt-BR')}`); } catch(_) {}
        } catch (err) {
          setStatus(status, '');
          if (debug) debug.textContent = 'Erro: ' + (err?.message || err);
          showToast('Erro ao processar cadastro: ' + (err?.message || err), 'danger');
        }

        loading && loading.classList.add('d-none');
        progress && progress.classList.add('d-none');
        cadastroBtn.classList.remove('btn-processing');
        cadastroInProgress = false;
        const progressBar = document.getElementById('cadastro_progressBar');
        if (progressBar) { progressBar.style.width = '0%'; progressBar.textContent = ''; progressBar.setAttribute('aria-valuenow','0'); }
      });
    }
  } catch (e) {
    /* noop */
  }
});
