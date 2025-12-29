document.addEventListener("DOMContentLoaded", function () {
  // Theme toggle (guarded: only run if element exists)
  const themeToggle = document.getElementById("themeToggle");
  if (themeToggle) {
    const setThemeHtml = (isDark) =>
      (themeToggle.innerHTML = isDark
        ? '<i class="bi bi-sun-fill animated-icon pulse" id="themeIcon"></i> Light'
        : '<i class="bi bi-moon-stars-fill animated-icon pulse" id="themeIcon"></i> Dark');
    if (localStorage.getItem("theme") === "dark") {
      document.body.classList.add("dark");
      setThemeHtml(true);
    } else {
      document.body.classList.remove("dark");
      setThemeHtml(false);
    }
    themeToggle.addEventListener("click", () => {
      document.body.classList.toggle("dark");
      const isDark = document.body.classList.contains("dark");
      setThemeHtml(isDark);
      localStorage.setItem("theme", isDark ? "dark" : "light");
      // When theme changes, update any raster icons in upload zones
      applyRasterInvertToUploadZones();
    });
  }

  // Tooltips
  const tooltipTriggerList = document.querySelectorAll(
    '[data-bs-toggle="tooltip"]'
  );
  [...tooltipTriggerList].map((el) => new bootstrap.Tooltip(el));

  // Helper functions
  /**
   * UI helper: set inner HTML status on an element (or clear if falsy)
   * @param {HTMLElement} elem
   * @param {string} html
   */
  function setStatus(elem, html) {
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
  /**
   * Apply or remove an invert filter to raster images inside upload zones
   * when dark mode is active. This is a pragmatic fallback for PNG/JPEG
   * icons that cannot be recolored via SVG/currentColor.
   */
  function applyRasterInvertToUploadZones() {
    try {
      const invert = document.body.classList.contains("dark");
      document.querySelectorAll(".upload_dropZone img").forEach((img) => {
        if (invert) img.classList.add("icon-invert-dark");
        else img.classList.remove("icon-invert-dark");
      });
    } catch (e) {
      // non-fatal
      console.debug("applyRasterInvertToUploadZones error", e);
    }
  }

  /**
   * Renderiza um resumo das estatísticas retornadas pelo endpoint de preview.
   * Comentários em português para facilitar manutenção futura.
   * O objetivo é mostrar rapidamente quantas correspondências foram
   * encontradas por método (CPF, exato, token, fuzzy, etc.).
   * Recebe um objeto `stats` que pode conter chaves numéricas ou objetos
   * mais complexos. Tentamos renderizar badges para valores simples e
   * um JSON legível para estruturas aninhadas.
   * @param {object} stats - objeto com estatísticas retornado pelo backend
   * @returns {string} HTML com as estatísticas formatadas
   */
  function renderPreviewStatsHtml(stats) {
    if (!stats || typeof stats !== "object") return "";
    const parts = [];
    // Percorre cada chave/valor das estatísticas
    for (const [k, v] of Object.entries(stats)) {
      const key = String(k || "").toLowerCase();
      // Mapeamento simples para rótulos legíveis em PT-BR
      const label = key.includes("cpf")
        ? "CPF"
        : key.includes("exact")
        ? "Exato"
        : key.includes("token")
        ? "Token"
        : key.includes("fuzzy")
        ? "Fuzzy"
        : k;
      // Valores primitivos são apresentados como badges
      if (
        v === null ||
        typeof v === "number" ||
        typeof v === "string" ||
        typeof v === "boolean"
      ) {
        parts.push(
          `<span class="badge bg-info text-dark me-1 preview-badge">${label}: ${v}</span>`
        );
      } else {
        // Para estruturas complexas (arrays/objetos) exibir um resumo compacto
        try {
          if (Array.isArray(v)) {
            parts.push(
              `<div class="mt-2"><strong>${label}:</strong> <span class="small text-muted">array[${v.length}]</span></div>`
            );
          } else if (typeof v === "object") {
            const keys = Object.keys(v || {})
              .slice(0, 5)
              .join(", ");
            parts.push(
              `<div class="mt-2"><strong>${label}:</strong> <span class="small text-muted">object{${keys}${
                Object.keys(v || {}).length > 5 ? ", ..." : ""
              }}</span></div>`
            );
          } else {
            parts.push(
              `<div class="mt-2"><strong>${label}:</strong> ${String(v)}</div>`
            );
          }
        } catch (e) {
          parts.push(
            `<div class="mt-2"><strong>${label}:</strong> ${String(v)}</div>`
          );
        }
      }
    }
    return `<div class="mt-3"><h6>Estatísticas</h6><div>${parts.join(
      ""
    )}</div></div>`;
  }

  /**
   * Gera HTML amigável para exibir correspondências que já estão INATIVAS
   * Espera um objeto onde cada chave é um tipo (ex: 'cpf','exact','token')
   * e o valor é um array de objetos (registros). Limita a 10 linhas por grupo.
   */
  function renderInactiveMatchesHtml(inactiveObj) {
    if (!inactiveObj || typeof inactiveObj !== "object") return "";
    const parts = [];
    // botão global para expandir/colapsar todos os grupos
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
      // descobrir colunas relevantes (priorizar CPF, NomeCompleto, Status, Email, Departamento, NomeEmpresa)
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
      const cols = [];
      // construir conjunto de chaves baseado nas primeiras linhas
      const keySet = new Set();
      sample.forEach((r) => Object.keys(r || {}).forEach((k) => keySet.add(k)));
      // ordenar chaves preferenciais primeiro
      preferred.forEach((k) => {
        if (keySet.has(k)) {
          cols.push(k);
          keySet.delete(k);
        }
      });
      // append remaining keys (limit to avoid huge tables)
      Array.from(keySet)
        .slice(0, 8)
        .forEach((k) => cols.push(k));

      // montar item do accordion
      const headingId = `inativosHeading${idx}`;
      const collapseId = `inativosCollapse${idx}`;
      let html = "";
      html += `<div class="accordion-item">`;
      html += `<h2 class="accordion-header" id="${headingId}">`;
      html += `<button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}" aria-expanded="false" aria-controls="${collapseId}">Inativos - ${group} <span class="ms-2 small text-muted">(${arr.length} total)</span></button>`;
      html += `</h2>`;
      html += `<div id="${collapseId}" class="accordion-collapse collapse" aria-labelledby="${headingId}" data-bs-parent="#inativosAccordion">`;
      html += `<div class="accordion-body">`;
      html += '<div class="inativos-table-wrapper">';
      html +=
        '<table class="table table-sm table-bordered"><thead class="table-light"><tr>' +
        cols.map((c) => `<th>${c}</th>`).join("") +
        "</tr></thead><tbody>";
      for (const row of sample) {
        html +=
          "<tr>" +
          cols
            .map((c) => `<td>${row && row[c] ? String(row[c]) : ""}</td>`)
            .join("") +
          "</tr>";
      }
      html += "</tbody></table>";
      if (arr.length > sample.length)
        html += `<div class="small text-muted">Mostrando ${sample.length} de ${arr.length} itens. Exporte para análise completa.</div>`;
      html += "</div>"; // wrapper
      html += `</div></div>`; // body + collapse
      html += `</div>`; // item
      parts.push(html);
      idx++;
    }
    parts.push("</div>"); // end accordion
    // Script que ativa os botões globais depois que o HTML é inserido no DOM
    parts.push(
      `<script>document.addEventListener('DOMContentLoaded', function(){ const expandBtn=document.getElementById('inativos_expand_all'); const collapseBtn=document.getElementById('inativos_collapse_all'); if(expandBtn){ expandBtn.addEventListener('click', ()=>{ document.querySelectorAll('#inativosAccordion .accordion-collapse').forEach(c=>{ if(!c.classList.contains('show')){ const bs = bootstrap.Collapse.getOrCreateInstance(c, {toggle:false}); bs.show(); }}); expandBtn.classList.add('d-none'); collapseBtn.classList.remove('d-none'); }); } if(collapseBtn){ collapseBtn.addEventListener('click', ()=>{ document.querySelectorAll('#inativosAccordion .accordion-collapse.show').forEach(c=>{ const bs = bootstrap.Collapse.getOrCreateInstance(c, {toggle:false}); bs.hide(); }); collapseBtn.classList.add('d-none'); expandBtn.classList.remove('d-none'); }); } });</script>`
    );
    return parts.join("");
  }
  async function postFiles(url, filesInput, extra = {}, single = false) {
    const fd = new FormData();
    if (single) {
      // filesInput.base should be a file input element
      if (
        !filesInput ||
        !filesInput.base ||
        !filesInput.base.files ||
        !filesInput.base.files[0]
      ) {
        throw new Error("Base file missing");
      }
      fd.append("base", filesInput.base.files[0]);
      // lista may be null when the user provided lista_text instead of an uploaded file
      if (
        filesInput.lista &&
        filesInput.lista.files &&
        filesInput.lista.files[0]
      ) {
        fd.append("lista", filesInput.lista.files[0]);
      }
    } else {
      for (const f of filesInput.files) fd.append("files[]", f);
    }
    for (const k in extra) fd.append(k, extra[k]);
    const xhr = new XMLHttpRequest();
    return new Promise((resolve, reject) => {
      xhr.upload.addEventListener("progress", (e) => {
        if (e.lengthComputable) {
          const percent = (e.loaded / e.total) * 100;
          const progressBar = url.includes("process_cadastro")
            ? document.getElementById("cadastro_progressBar")
            : document.getElementById("inativacao_progressBar");
          progressBar.style.width = `${percent}%`;
          progressBar.setAttribute("aria-valuenow", percent);
          progressBar.textContent = `${percent.toFixed(1)}%`;
        }
      });
      // If a global API base is defined (set by index.html), prefix relative URLs
      const fullUrl =
        window.API_BASE && url.startsWith("/") ? window.API_BASE + url : url;
      xhr.open("POST", fullUrl, true);
      xhr.responseType = "blob";
      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve({ blob: xhr.response });
          return;
        }
        // Em caso de erro, tentar extrair mensagem do corpo (JSON ou texto)
        const blob = xhr.response;
        if (blob) {
          const reader = new FileReader();
          reader.onload = () => {
            const text = reader.result || "";
            try {
              const obj = JSON.parse(text);
              const msg = obj.error || obj.message || text;
              // If there's an 'errors' object, append it for more detail
              const details = obj.errors
                ? ` | details: ${JSON.stringify(obj.errors)}`
                : "";
              reject(new Error(`Server ${xhr.status}: ${msg}${details}`));
            } catch (e) {
              // não JSON
              reject(
                new Error(`Server ${xhr.status}: ${text || xhr.statusText}`)
              );
            }
          };
          reader.onerror = () =>
            reject(new Error(`Server ${xhr.status}: ${xhr.statusText}`));
          reader.readAsText(blob);
        } else {
          reject(new Error(`Server ${xhr.status}: ${xhr.statusText}`));
        }
      };
      xhr.onerror = () => reject(new Error("Erro de rede"));
      xhr.send(fd);
    });
  }
  // Envia FormData e espera JSON de resposta (usado para preview)
  async function postFormDataJson(url, formData) {
    const fullUrl =
      window.API_BASE && url.startsWith("/") ? window.API_BASE + url : url;
    const resp = await fetch(fullUrl, { method: "POST", body: formData });
    const text = await resp.text();
    try {
      const obj = text ? JSON.parse(text) : {};
      if (!resp.ok) {
        const msg = obj.error || obj.message || resp.statusText;
        const details = obj.errors
          ? ` | details: ${JSON.stringify(obj.errors)}`
          : "";
        throw new Error(`Server ${resp.status}: ${msg}${details}`);
      }
      return obj;
    } catch (e) {
      // se não for JSON
      if (!resp.ok)
        throw new Error(`Server ${resp.status}: ${text || resp.statusText}`);
      return {};
    }
  }
  /**
   * showToast: exibe uma notificação amigável usando Bootstrap Toast
   * Comentários em português: aceitar tipos 'success', 'info', 'error'
   * Garantir aria-live para leitores de tela.
   */
  function showToast(message, type = "success") {
    const toastContainer = document.getElementById("toastContainer");
    if (!toastContainer) return console.warn("toastContainer não encontrado");
    // mapa simples de tipos para classes bootstrap
    const bg =
      type === "success" ? "success" : type === "info" ? "info" : "danger";
    const toast = document.createElement("div");
    toast.setAttribute("role", "status");
    toast.setAttribute("aria-live", "polite");
    toast.className = `toast align-items-center text-white bg-${bg} border-0`;
    toast.innerHTML = `<div class="d-flex"><div class="toast-body">${message}</div><button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Fechar"></button></div>`;
    toastContainer.appendChild(toast);
    const bsToast = new bootstrap.Toast(toast, { delay: 3500 });
    bsToast.show();
    toast.addEventListener("hidden.bs.toast", () => toast.remove());
  }

  // API Status
  const apiStatusBtn = document.getElementById("apiStatusBtn");
  const apiStatusIcon = apiStatusBtn?.querySelector("i");
  if (apiStatusBtn) {
    apiStatusBtn.addEventListener("click", async () => {
      apiStatusIcon?.classList.add("spin-rotating");
      const base = window.API_BASE ? window.API_BASE.replace(/\/$/, "") : "";
      const urls = [
        `${base}/health`,
        `${base}/api/health`,
        base ? `${base}/` : "/",
      ];
      async function tryGet(u) {
        try {
          const r = await fetch(u, { method: "GET" });
          return r.ok;
        } catch (_) {
          return false;
        }
      }
      let ok = false;
      for (const u of urls) {
        /* eslint-disable no-await-in-loop */
        if (await tryGet(u)) { ok = true; break; }
      }
      showToast(ok ? "API Online" : "API Offline", ok ? "success" : "danger");
      apiStatusIcon?.classList.remove("spin-rotating");
    });
  }

  // Run raster invert at startup based on current theme
  applyRasterInvertToUploadZones();

  // Análise - Drag-and-Drop
  const analiseUploadArea = document.getElementById("analise_uploadArea");
  const analiseFiles = document.getElementById("analise_files");
  const analiseFileFeedback = document.getElementById("analise_fileFeedback");
  if (analiseUploadArea && analiseFiles && analiseFileFeedback) {
    // Tornar a área acessível via teclado: Enter ou Espaço abrem o seletor de arquivos
    analiseUploadArea.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        analiseFiles.click();
      }
    });
    analiseUploadArea.addEventListener("dragover", (e) => {
      e.preventDefault();
      analiseUploadArea.classList.add("dragover");
    });
    analiseUploadArea.addEventListener("dragleave", () =>
      analiseUploadArea.classList.remove("dragover")
    );
    analiseUploadArea.addEventListener("drop", (e) => {
      e.preventDefault();
      analiseUploadArea.classList.remove("dragover");
      analiseFiles.files = e.dataTransfer.files;
      const file = analiseFiles.files[0];
      analiseFileFeedback.classList.remove("d-none");
      analiseFileFeedback.innerHTML = `<i class="bi bi-check-circle-fill animated-icon pulse me-2"></i>Anexo bem-sucedido: ${file.name}`;
      showToast("Arquivo anexado com sucesso!", "success");
      addToHistory(`Análise: ${file.name} - ${new Date().toLocaleString()}`);
      processAnaliseFile(file);
    });
    analiseFiles.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (file) {
        analiseFileFeedback.classList.remove("d-none");
        analiseFileFeedback.innerHTML = `<i class="bi bi-check-circle-fill animated-icon pulse me-2"></i>Anexo bem-sucedido: ${file.name}`;
        showToast("Arquivo anexado com sucesso!", "success");
        addToHistory(`Análise: ${file.name} - ${new Date().toLocaleString()}`);
        processAnaliseFile(file);
      }
    });
  }

  function processAnaliseFile(file) {
    const loading = document.getElementById("analise_loading");
    const tableContainer = document.getElementById("analise_tableContainer");
    const tableHead = document.getElementById("analise_tableHead");
    const tableBody = document.getElementById("analise_tableBody");
    const summaryContainer = document.getElementById(
      "analise_summaryContainer"
    );
    const summaryText = document.getElementById("analise_summaryText");
    const searchInput = document.getElementById("analise_searchInput");
    const debug = document.getElementById("analise_debug");
    const toggleCompactView = document.getElementById("toggleCompactView");

    if (!file || !file.name.match(/\.(xlsx|xls)$/)) {
      showToast("Por favor, envie um arquivo Excel (.xlsx ou .xls).", "danger");
      return;
    }

    loading.classList.remove("d-none");
    tableContainer.classList.add("d-none");
    summaryContainer.classList.add("d-none");
    searchInput.classList.add("d-none");
    debug.textContent = "";

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (jsonData.length < 2) {
          showToast(
            "O arquivo Excel contém apenas cabeçalhos, sem dados preenchidos.",
            "info"
          );
          loading.classList.add("d-none");
          return;
        }

        const headers = jsonData[0].map((header) => header.trim());
        const dataRows = jsonData
          .slice(1)
          .map((row) =>
            headers.reduce((obj, header, index) => {
              obj[header] =
                row[index] !== undefined
                  ? row[index].toString().trim().toUpperCase()
                  : "";
              return obj;
            }, {})
          )
          .filter((row) => Object.values(row).some((val) => val !== ""));

        summaryText.textContent = `Total de cadastros: ${dataRows.length}`;
        summaryContainer.classList.remove("d-none");

        tableHead.innerHTML =
          "<tr>" +
          headers.map((header) => `<th>${header}</th>`).join("") +
          "</tr>";
        tableBody.innerHTML = dataRows
          .slice(0, 10)
          .map(
            (row) =>
              `<tr>${headers
                .map((header) => `<td>${row[header] || "N/A"}</td>`)
                .join("")}</tr>`
          )
          .join("");

        searchInput.classList.remove("d-none");
        searchInput.addEventListener("input", (e) => {
          const term = e.target.value.toLowerCase();
          tableBody.innerHTML = dataRows
            .filter((row) =>
              Object.values(row).some((val) => val.toLowerCase().includes(term))
            )
            .slice(0, 10)
            .map(
              (row) =>
                `<tr>${headers
                  .map((header) => `<td>${row[header] || "N/A"}</td>`)
                  .join("")}</tr>`
            )
            .join("");
        });

        toggleCompactView.addEventListener("click", () => {
          document.body.classList.toggle("compact");
          toggleCompactView.textContent = document.body.classList.contains(
            "compact"
          )
            ? "Modo Padrão"
            : "Modo Compacto";
          showToast("Visualização alterada!", "success");
        });

        tableContainer.classList.remove("d-none");
        loading.classList.add("d-none");
        showToast("Ficha carregada com sucesso!", "success");
      } catch (error) {
        loading.classList.add("d-none");
        debug.textContent = "Erro: " + error.message;
        showToast("Erro ao processar o arquivo: " + error.message, "danger");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // Cadastro - Drag-and-Drop
  const cadastroUploadArea = document.getElementById("cadastro_uploadArea");
  const cadastroFiles = document.getElementById("cadastro_files");
  if (cadastroUploadArea && cadastroFiles) {
    // Suporte a ativação por teclado
    cadastroUploadArea.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        cadastroFiles.click();
      }
    });
    cadastroUploadArea.addEventListener("dragover", (e) => {
      e.preventDefault();
      cadastroUploadArea.classList.add("dragover");
    });
    cadastroUploadArea.addEventListener("dragleave", () =>
      cadastroUploadArea.classList.remove("dragover")
    );
    cadastroUploadArea.addEventListener("drop", (e) => {
      e.preventDefault();
      cadastroUploadArea.classList.remove("dragover");
      cadastroFiles.files = e.dataTransfer.files;
      validateAndShowFiles(cadastroFiles);
      addToHistory(
        `Cadastro: ${
          e.dataTransfer.files[0]?.name
        } - ${new Date().toLocaleString()}`
      );
    });
  }

  function validateAndShowFiles(filesInput) {
    const maxFiles = 5;
    const maxSize = 10 * 1024 * 1024;
    if (filesInput.files.length > maxFiles) {
      showToast(`Máximo de ${maxFiles} arquivos excedido.`, "danger");
      filesInput.value = "";
      return;
    }
    for (let file of filesInput.files) {
      if (file.size > maxSize) {
        showToast(`Arquivo ${file.name} excede 10MB.`, "danger");
        filesInput.value = "";
        return;
      }
    }
    document.getElementById("cadastro_progress").classList.remove("d-none");
    // show selected file names
    try {
      const display = document.getElementById("cadastro_fileNames");
      if (display)
        display.textContent = Array.from(filesInput.files)
          .map((f) => f.name)
          .join(", ");
    } catch (e) {
      /* no-op */
    }
  }

  // Cadastro - Button click
  const cadastroBtn = document.getElementById("cadastro_btn");
  // flag para evitar envios duplicados
  let cadastroInProgress = false;
  if (cadastroBtn) {
    cadastroBtn.addEventListener("click", async () => {
      // se já estiver processando, ignorar clique
      if (cadastroInProgress) return;
      cadastroInProgress = true;
      cadastroBtn.classList.add("btn-processing");
      const files = document.getElementById("cadastro_files");
      const status = document.getElementById("cadastro_status");
      const debug = document.getElementById("cadastro_debug");
      const loading = document.getElementById("cadastro_loading");
      const progress = document.getElementById("cadastro_progress");
      const loginChoice = document.getElementById(
        "cadastro_login_choice"
      ).value;
      const fluxo = document.getElementById("cadastro_fluxo").value;

      if (!status || !debug || !loading || !progress)
        console.error("Elementos de status não encontrados!");
      setStatus(
        status,
        '<span class="spinner-border spinner-border-sm text-primary me-2" role="status"></span>Processando...'
      );
      debug.textContent = "";
      loading.classList.remove("d-none");
      progress.classList.remove("d-none");

      try {
        if (!files.files.length)
          throw new Error("Selecione pelo menos um arquivo");
        const extraData = { login_choice: loginChoice, fluxo: fluxo };
        if (fluxo === "SELF") {
          extraData.vip = "N";
          extraData.viajanteMasterNacional = "N";
          extraData.viajanteMasterInternacional = "N";
          extraData.solicitanteMaster = "N";
          extraData.masterAdiantamento = "N";
          extraData.masterReembolso = "N";
        } else if (fluxo === "FRONT") {
          extraData.vip = "N";
          extraData.viajanteMasterNacional = "S";
          extraData.viajanteMasterInternacional = "S";
          extraData.solicitanteMaster = "N";
          extraData.masterAdiantamento = "N";
          extraData.masterReembolso = "N";
        }
        const resp = await postFiles("/api/process_cadastro", files, extraData);
        if (!resp.blob || resp.blob.size === 0)
          throw new Error("Arquivo gerado inválido");
        downloadBlob(resp.blob, "saida_cadastro.xlsx");
        setStatus(status, '<span class="text-success">✔ Concluído</span>');
        showToast(
          `Cadastro processado! Opções: ${loginChoice}, ${fluxo}.`,
          "success"
        );
        addToHistory(
          `Cadastro gerado: ${
            files.files[0].name
          } - ${new Date().toLocaleString()}`
        );
      } catch (err) {
        setStatus(status, "");
        debug.textContent = "Erro: " + (err.message || err);
        showToast(
          "Erro ao processar cadastro: " + (err.message || err),
          "danger"
        );
      }
      loading.classList.add("d-none");
      progress.classList.add("d-none");
      cadastroBtn.classList.remove("btn-processing");
      cadastroInProgress = false;
      document.getElementById("cadastro_progressBar").style.width = "0%";
      document.getElementById("cadastro_progressBar").textContent = "";
    });
  }

  // Inativação - Drag-and-Drop
  const inativacaoBaseUploadArea = document.getElementById(
    "inativacao_base_uploadArea"
  );
  const inativacaoBase = document.getElementById("inativacao_base");
  if (inativacaoBaseUploadArea && inativacaoBase) {
    // ativação por teclado na zona de upload da base
    inativacaoBaseUploadArea.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        inativacaoBase.click();
      }
    });
    inativacaoBaseUploadArea.addEventListener("dragover", (e) => {
      e.preventDefault();
      inativacaoBaseUploadArea.classList.add("dragover");
    });
    inativacaoBaseUploadArea.addEventListener("dragleave", () =>
      inativacaoBaseUploadArea.classList.remove("dragover")
    );
    inativacaoBaseUploadArea.addEventListener("drop", (e) => {
      e.preventDefault();
      inativacaoBaseUploadArea.classList.remove("dragover");
      inativacaoBase.files = e.dataTransfer.files;
      addToHistory(
        `Inativação - Base: ${
          e.dataTransfer.files[0]?.name
        } - ${new Date().toLocaleString()}`
      );
    });
  }

  const inativacaoListaUploadArea = document.getElementById(
    "inativacao_lista_uploadArea"
  );
  const inativacaoLista = document.getElementById("inativacao_lista");
  if (inativacaoListaUploadArea && inativacaoLista) {
    // ativação por teclado na zona de upload da lista
    inativacaoListaUploadArea.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        inativacaoLista.click();
      }
    });
    inativacaoListaUploadArea.addEventListener("dragover", (e) => {
      e.preventDefault();
      inativacaoListaUploadArea.classList.add("dragover");
    });
    inativacaoListaUploadArea.addEventListener("dragleave", () =>
      inativacaoListaUploadArea.classList.remove("dragover")
    );
    inativacaoListaUploadArea.addEventListener("drop", (e) => {
      e.preventDefault();
      inativacaoListaUploadArea.classList.remove("dragover");
      inativacaoLista.files = e.dataTransfer.files;
      addToHistory(
        `Inativação - Lista: ${
          e.dataTransfer.files[0]?.name
        } - ${new Date().toLocaleString()}`
      );
    });
  }
  // Inativação - Button click
  const inativacaoBtn = document.getElementById("inativacao_btn");
  // flag para evitar envios duplicados
  let inativacaoInProgress = false;
  if (inativacaoBtn) {
    inativacaoBtn.addEventListener("click", async () => {
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
          throw new Error(errPreview.message || "Erro no preview");
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

        let infoHtml = `<p>Foram encontradas <strong>${displayCount}</strong> correspondência(s).</p>`;
        if (sample && sample.length) {
          const keys = Object.keys(sample[0]).slice(0, 6);
          const rowsHtml = sample
            .slice(0, 10)
            .map((row) => {
              return `<tr>${keys
                .map((k) => `<td>${row[k] || ""}</td>`)
                .join("")}</tr>`;
            })
            .join("");
          infoHtml += `
            <div style="max-height:200px;overflow:auto;margin-top:8px;">
              <table class="table table-sm table-bordered"><thead><tr>
                ${keys.map((k) => `<th>${k}</th>`).join("")}
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
          const headerHtml = colsToShow.map((c) => `<th>${c}</th>`).join("");
          const bodyHtml = initialRows
            .map(
              (row) =>
                `<tr>${colsToShow
                  .map((c) => `<td>${row[c] ?? ""}</td>`)
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

        const confirmed = await Swal.fire({
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
        debug.textContent = "Erro: " + (err.message || err);
        showToast(
          "Erro ao processar inativação: " + (err.message || err),
          "danger"
        );
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

  // Histórico
  const historicoList = document.getElementById("historico_list");
  const clearHistoryBtn = document.getElementById("clearHistoryBtn");
  function addToHistory(action) {
    const history = JSON.parse(localStorage.getItem("history") || "[]");
    history.push(action);
    localStorage.setItem("history", JSON.stringify(history.slice(-10))); // Limita a 10 itens
    renderHistory();
  }
  function renderHistory() {
    const history = JSON.parse(localStorage.getItem("history") || "[]");
    historicoList.innerHTML = history
      .map(
        (item, index) =>
          `<li class="list-group-item">${index + 1}. ${item}</li>`
      )
      .join("");
  }
  if (clearHistoryBtn) {
    clearHistoryBtn.addEventListener("click", () => {
      localStorage.removeItem("history");
      renderHistory();
      showToast("Histórico limpo!", "success");
    });
  }
  renderHistory();
});
