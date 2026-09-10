import { useMemo, useRef, useState } from "react";
import { Modal } from "../../components/Modal";
import { formatCpf, useInativacao } from "./useInativacao";

const PAGE_SIZE = 10;

function rowStatusClass(found: boolean, status: string): string {
  const s = (status || "").trim().toUpperCase();
  if (!found) return "bg-warning/15";
  if (s === "ATIVO") return "bg-success/10";
  if (s) return "bg-danger/10";
  return "";
}

export function InativacaoTab() {
  const inativacao = useInativacao();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);
  const [resultSearch, setResultSearch] = useState("");
  const [page, setPage] = useState(1);
  const [helpOpen, setHelpOpen] = useState(false);

  const filtered = useMemo(() => {
    const term = resultSearch.trim().toLowerCase();
    if (!term) return inativacao.results;
    return inativacao.results.filter(
      (r) => r.nome.toLowerCase().includes(term) || r.cpf.includes(term),
    );
  }, [inativacao.results, resultSearch]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const pageRows = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  function pickFile(file: File | null) {
    inativacao.setBaseFile(file);
    setPage(1);
  }

  function clearBase() {
    pickFile(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
    inativacao.setResults([]);
  }

  async function handleSearch() {
    await inativacao.search();
    setPage(1);
  }

  async function handleGenerate() {
    await inativacao.generate(clearBase);
  }

  const showResults = inativacao.results.length > 0 || inativacao.searching;

  return (
    <div className="rounded-xl border border-black/10 bg-surface p-4 shadow-sm dark:border-white/10 dark:bg-surface-dark sm:p-6">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <h2 className="text-lg font-semibold text-slate-900 dark:text-slate-100">Inativação</h2>
        <button
          id="inativacao_help_btn"
          type="button"
          aria-label="Como usar a inativação"
          title="Guia da inativação"
          onClick={() => setHelpOpen(true)}
          className="rounded-lg border border-info px-3 py-1.5 text-sm font-medium text-info hover:bg-info/10"
        >
          Como usar
        </button>
      </div>

      <Modal open={helpOpen} onClose={() => setHelpOpen(false)} title="Inativação">
        <p className="mb-2 font-semibold">Passos para inativação rápida</p>
        <ol className="list-decimal space-y-1 pl-5">
          <li>
            <strong>1º Passo:</strong> envie a base de usuários em Excel (.xlsx).
          </li>
          <li>
            <strong>2º Passo:</strong> cole <strong>CPFs</strong>, <strong>nomes completos</strong> ou <strong>e-mails</strong>, um por linha.
          </li>
          <li>
            <strong>3º Passo:</strong> use <em>Buscar</em> para conferir quem será afetado e os status atuais.
          </li>
          <li>
            <strong>4º Passo:</strong> finalize em <strong>Gerar</strong> para baixar o relatório da inativação.
          </li>
        </ol>
      </Modal>

      <div
        id="inativacao_base_uploadArea"
        aria-live="polite"
        tabIndex={0}
        role="button"
        aria-label="Upload da base de inativacao. Pressione para selecionar arquivo"
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            fileInputRef.current?.click();
          }
        }}
        onDragOver={(e) => {
          e.preventDefault();
          setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragOver(false);
          const file = e.dataTransfer.files?.[0];
          if (file) pickFile(file);
        }}
        className={`mb-4 rounded-lg border-2 border-dashed p-6 text-center transition-colors ${
          dragOver ? "border-accent bg-accent/5" : "border-accent/40"
        } dark:border-accent-dark/40`}
      >
        <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">
          Arraste a base (.xlsx) ou selecione o arquivo.
        </p>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          id="inativacao_base"
          aria-label="Selecionar base para inativação"
          className="hidden"
          onChange={(e) => pickFile(e.target.files?.[0] ?? null)}
        />
        <button
          type="button"
          title="Selecionar planilha base de usuários"
          onClick={() => fileInputRef.current?.click()}
          className="rounded-lg bg-accent px-4 py-2 text-sm font-medium text-white hover:bg-accent-hover"
        >
          Selecionar
        </button>
        <div id="inativacao_base_feedback" aria-live="polite" className="mt-3">
          {inativacao.baseFile && (
            <span className="inline-flex items-center gap-2 rounded-full bg-success/10 px-3 py-1 text-sm text-success">
              {inativacao.baseFile.name}
              <button
                id="inativacao_clear_base_btn"
                type="button"
                title="Remover base"
                aria-label="Limpar base enviada"
                onClick={(e) => {
                  e.stopPropagation();
                  clearBase();
                }}
                className="text-success hover:text-danger"
              >
                ✕
              </button>
            </span>
          )}
        </div>
      </div>

      <div className="mt-3">
        <label htmlFor="lista_text" className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-200">
          Informe abaixo como deseja localizar os usuários.
        </label>
        <textarea
          id="lista_text"
          rows={4}
          placeholder={"Ex: João Silva\n12345678901\nusuario@example.com"}
          aria-describedby="lista_valid_summary lista_duplicates_warning"
          title="Insira (nomes, CPFs ou e-mails) e eles serão validados automaticamente"
          value={inativacao.listText}
          onChange={(e) => inativacao.setListText(e.target.value)}
          className="w-full rounded-lg border border-black/10 bg-white p-3 text-sm text-slate-900 outline-none focus:border-accent focus:ring-2 focus:ring-accent/25 dark:border-white/10 dark:bg-surface-dark-alt dark:text-slate-100"
        />
        <div id="lista_valid_summary" className="mt-2 text-sm text-slate-500 dark:text-slate-400">
          <span className="rounded bg-accent/10 px-2 py-0.5 text-accent dark:text-accent-dark">
            {inativacao.classification.totalValid}
          </span>{" "}
          itens válidos (CPF, Nome Completo ou E-mail)
        </div>
        {inativacao.classification.duplicates.length > 0 && (
          <div id="lista_duplicates_warning" className="mt-1 text-sm text-warning">
            CPFs duplicados: {inativacao.classification.duplicates.join(", ")}
          </div>
        )}
      </div>

      {inativacao.generating && (
        <div id="inativacao_progress" className="mt-3 h-2 overflow-hidden rounded-full bg-black/10 dark:bg-white/10">
          <div
            id="inativacao_progressBar"
            role="progressbar"
            aria-valuenow={Math.round(inativacao.progress)}
            aria-valuemin={0}
            aria-valuemax={100}
            style={{ width: `${inativacao.progress}%` }}
            className="h-full bg-accent transition-[width]"
          />
        </div>
      )}

      <button
        id="inativacao_btn"
        aria-label="Buscar usuários para inativação"
        title="Executar a busca de itens digitados na base"
        disabled={!inativacao.canSearch || inativacao.searching}
        onClick={handleSearch}
        className="mt-4 rounded-lg bg-accent px-4 py-2 text-sm font-medium text-white hover:bg-accent-hover disabled:cursor-not-allowed disabled:opacity-50"
      >
        {inativacao.searching ? "Buscando..." : "Buscar Usuários"}
      </button>

      <div id="inativacao_status" className="mt-3 text-sm text-slate-500 dark:text-slate-400" aria-live="polite" />
      {inativacao.debugMsg && (
        <div id="inativacao_debug" className="mt-3 text-sm text-danger" aria-live="assertive">
          {inativacao.debugMsg}
        </div>
      )}

      {showResults && (
        <>
          <div id="inativacao_results_controls" className="mt-4 flex flex-wrap items-center gap-2">
            <input
              id="result_search"
              placeholder="Filtrar por nome ou CPF"
              aria-label="Filtrar resultados"
              value={resultSearch}
              onChange={(e) => {
                setResultSearch(e.target.value);
                setPage(1);
              }}
              className="min-w-[220px] flex-1 rounded-lg border border-black/10 bg-white px-3 py-1.5 text-sm text-slate-900 outline-none focus:border-accent focus:ring-2 focus:ring-accent/25 dark:border-white/10 dark:bg-surface-dark-alt dark:text-slate-100"
            />
            {inativacao.results.length > 0 && (
              <button
                id="inativacao_generate_btn"
                aria-label="Gerar inativação"
                title="Processar e baixar o relatório final"
                disabled={inativacao.generating}
                onClick={handleGenerate}
                className="rounded-lg bg-success px-4 py-2 text-sm font-medium text-white hover:brightness-95 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {inativacao.generating ? "Gerando..." : "Gerar Inativação"}
              </button>
            )}
          </div>

          <div id="inativacao_results" className="mt-3 overflow-x-auto rounded-lg border border-black/10 dark:border-white/10">
            <table id="results_table" className="w-full text-sm">
              <thead className="bg-black/[.03] dark:bg-white/[.04]">
                <tr>
                  <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">
                    Nome Completo
                  </th>
                  <th scope="col" className="whitespace-nowrap px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">
                    CPF
                  </th>
                  <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">
                    E-mail
                  </th>
                  <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">
                    Status Atual
                  </th>
                </tr>
              </thead>
              <tbody id="results_body">
                {pageRows.map((r, i) => (
                  <tr
                    key={`${r.cpf}-${i}`}
                    data-cpf={r.cpf}
                    className={`border-t border-black/5 dark:border-white/5 ${rowStatusClass(r.found, r.status_atual)}`}
                  >
                    <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{r.nome}</td>
                    <td className="px-3 py-1.5 font-mono text-slate-800 dark:text-slate-100">{formatCpf(r.cpf)}</td>
                    <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{r.email}</td>
                    <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{r.status_atual}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div id="results_pagination" className="mt-2 flex items-center justify-between">
            <button
              id="page_prev"
              disabled={page <= 1}
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              className="rounded-lg border border-black/15 px-3 py-1 text-sm disabled:cursor-not-allowed disabled:opacity-40 dark:border-white/15 dark:text-slate-200"
            >
              Anterior
            </button>
            <span id="page_info" className="text-sm text-slate-500 dark:text-slate-400">
              Página {page} de {totalPages}
            </span>
            <button
              id="page_next"
              disabled={page >= totalPages}
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              className="rounded-lg border border-black/15 px-3 py-1 text-sm disabled:cursor-not-allowed disabled:opacity-40 dark:border-white/15 dark:text-slate-200"
            >
              Próxima
            </button>
          </div>
        </>
      )}
    </div>
  );
}
