import { useMemo, useRef, useState } from "react";
import { Modal } from "../../components/Modal";
import { formatCpf, useInativacao } from "./useInativacao";

const PAGE_SIZE = 10;

function rowStatusClass(found: boolean, status: string): string {
  const s = (status || "").trim().toUpperCase();
  if (!found) return "tw:bg-warning/15";
  if (s === "ATIVO") return "tw:bg-success/10";
  if (s) return "tw:bg-danger/10";
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
    <div className="tw:rounded-xl tw:border tw:border-black/10 tw:bg-surface tw:p-4 tw:shadow-sm tw:dark:border-white/10 tw:dark:bg-surface-dark tw:sm:p-6">
      <div className="tw:mb-3 tw:flex tw:flex-wrap tw:items-center tw:justify-between tw:gap-2">
        <h2 className="tw:text-lg tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">Inativação</h2>
        <button
          id="inativacao_help_btn"
          type="button"
          aria-label="Como usar a inativação"
          title="Guia da inativação"
          onClick={() => setHelpOpen(true)}
          className="tw:rounded-lg tw:border tw:border-info tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-info tw:hover:bg-info/10"
        >
          Como usar
        </button>
      </div>

      <Modal open={helpOpen} onClose={() => setHelpOpen(false)} title="Inativação">
        <p className="tw:mb-2 tw:font-semibold">Passos para inativação rápida</p>
        <ol className="tw:list-decimal tw:space-y-1 tw:pl-5">
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
        className={`tw:mb-4 tw:rounded-lg tw:border-2 tw:border-dashed tw:p-6 tw:text-center tw:transition-colors ${
          dragOver ? "tw:border-accent tw:bg-accent/5" : "tw:border-accent/40"
        } tw:dark:border-accent-dark/40`}
      >
        <p className="tw:mb-3 tw:text-sm tw:text-slate-500 tw:dark:text-slate-400">
          Arraste a base (.xlsx) ou selecione o arquivo.
        </p>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          id="inativacao_base"
          aria-label="Selecionar base para inativação"
          className="tw:hidden"
          onChange={(e) => pickFile(e.target.files?.[0] ?? null)}
        />
        <button
          type="button"
          title="Selecionar planilha base de usuários"
          onClick={() => fileInputRef.current?.click()}
          className="tw:rounded-lg tw:bg-accent tw:px-4 tw:py-2 tw:text-sm tw:font-medium tw:text-white tw:hover:bg-accent-hover"
        >
          Selecionar
        </button>
        <div id="inativacao_base_feedback" aria-live="polite" className="tw:mt-3">
          {inativacao.baseFile && (
            <span className="tw:inline-flex tw:items-center tw:gap-2 tw:rounded-full tw:bg-success/10 tw:px-3 tw:py-1 tw:text-sm tw:text-success">
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
                className="tw:text-success tw:hover:text-danger"
              >
                ✕
              </button>
            </span>
          )}
        </div>
      </div>

      <div className="tw:mt-3">
        <label htmlFor="lista_text" className="tw:mb-1 tw:block tw:text-sm tw:font-medium tw:text-slate-700 tw:dark:text-slate-200">
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
          className="tw:w-full tw:rounded-lg tw:border tw:border-black/10 tw:bg-white tw:p-3 tw:text-sm tw:text-slate-900 tw:outline-none tw:focus:border-accent tw:focus:ring-2 tw:focus:ring-accent/25 tw:dark:border-white/10 tw:dark:bg-surface-dark-alt tw:dark:text-slate-100"
        />
        <div id="lista_valid_summary" className="tw:mt-2 tw:text-sm tw:text-slate-500 tw:dark:text-slate-400">
          <span className="tw:rounded tw:bg-accent/10 tw:px-2 tw:py-0.5 tw:text-accent tw:dark:text-accent-dark">
            {inativacao.classification.totalValid}
          </span>{" "}
          itens válidos (CPF, Nome Completo ou E-mail)
        </div>
        {inativacao.classification.duplicates.length > 0 && (
          <div id="lista_duplicates_warning" className="tw:mt-1 tw:text-sm tw:text-warning">
            CPFs duplicados: {inativacao.classification.duplicates.join(", ")}
          </div>
        )}
      </div>

      {inativacao.generating && (
        <div id="inativacao_progress" className="tw:mt-3 tw:h-2 tw:overflow-hidden tw:rounded-full tw:bg-black/10 tw:dark:bg-white/10">
          <div
            id="inativacao_progressBar"
            role="progressbar"
            aria-valuenow={Math.round(inativacao.progress)}
            aria-valuemin={0}
            aria-valuemax={100}
            style={{ width: `${inativacao.progress}%` }}
            className="tw:h-full tw:bg-accent tw:transition-[width]"
          />
        </div>
      )}

      <button
        id="inativacao_btn"
        aria-label="Buscar usuários para inativação"
        title="Executar a busca de itens digitados na base"
        disabled={!inativacao.canSearch || inativacao.searching}
        onClick={handleSearch}
        className="tw:mt-4 tw:rounded-lg tw:bg-accent tw:px-4 tw:py-2 tw:text-sm tw:font-medium tw:text-white tw:hover:bg-accent-hover tw:disabled:cursor-not-allowed tw:disabled:opacity-50"
      >
        {inativacao.searching ? "Buscando..." : "Buscar Usuários"}
      </button>

      <div id="inativacao_status" className="tw:mt-3 tw:text-sm tw:text-slate-500 tw:dark:text-slate-400" aria-live="polite" />
      {inativacao.debugMsg && (
        <div id="inativacao_debug" className="tw:mt-3 tw:text-sm tw:text-danger" aria-live="assertive">
          {inativacao.debugMsg}
        </div>
      )}

      {showResults && (
        <>
          <div id="inativacao_results_controls" className="tw:mt-4 tw:flex tw:flex-wrap tw:items-center tw:gap-2">
            <input
              id="result_search"
              placeholder="Filtrar por nome ou CPF"
              aria-label="Filtrar resultados"
              value={resultSearch}
              onChange={(e) => {
                setResultSearch(e.target.value);
                setPage(1);
              }}
              className="tw:min-w-[220px] tw:flex-1 tw:rounded-lg tw:border tw:border-black/10 tw:bg-white tw:px-3 tw:py-1.5 tw:text-sm tw:text-slate-900 tw:outline-none tw:focus:border-accent tw:focus:ring-2 tw:focus:ring-accent/25 tw:dark:border-white/10 tw:dark:bg-surface-dark-alt tw:dark:text-slate-100"
            />
            {inativacao.results.length > 0 && (
              <button
                id="inativacao_generate_btn"
                aria-label="Gerar inativação"
                title="Processar e baixar o relatório final"
                disabled={inativacao.generating}
                onClick={handleGenerate}
                className="tw:rounded-lg tw:bg-success tw:px-4 tw:py-2 tw:text-sm tw:font-medium tw:text-white tw:hover:brightness-95 tw:disabled:cursor-not-allowed tw:disabled:opacity-50"
              >
                {inativacao.generating ? "Gerando..." : "Gerar Inativação"}
              </button>
            )}
          </div>

          <div id="inativacao_results" className="tw:mt-3 tw:overflow-x-auto tw:rounded-lg tw:border tw:border-black/10 tw:dark:border-white/10">
            <table id="results_table" className="tw:w-full tw:text-sm">
              <thead className="tw:bg-black/[.03] tw:dark:bg-white/[.04]">
                <tr>
                  <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">
                    Nome Completo
                  </th>
                  <th scope="col" className="tw:whitespace-nowrap tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">
                    CPF
                  </th>
                  <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">
                    E-mail
                  </th>
                  <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">
                    Status Atual
                  </th>
                </tr>
              </thead>
              <tbody id="results_body">
                {pageRows.map((r, i) => (
                  <tr
                    key={`${r.cpf}-${i}`}
                    data-cpf={r.cpf}
                    className={`tw:border-t tw:border-black/5 tw:dark:border-white/5 ${rowStatusClass(r.found, r.status_atual)}`}
                  >
                    <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{r.nome}</td>
                    <td className="tw:px-3 tw:py-1.5 tw:font-mono tw:text-slate-800 tw:dark:text-slate-100">{formatCpf(r.cpf)}</td>
                    <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{r.email}</td>
                    <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{r.status_atual}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div id="results_pagination" className="tw:mt-2 tw:flex tw:items-center tw:justify-between">
            <button
              id="page_prev"
              disabled={page <= 1}
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              className="tw:rounded-lg tw:border tw:border-black/15 tw:px-3 tw:py-1 tw:text-sm tw:disabled:cursor-not-allowed tw:disabled:opacity-40 tw:dark:border-white/15 tw:dark:text-slate-200"
            >
              Anterior
            </button>
            <span id="page_info" className="tw:text-sm tw:text-slate-500 tw:dark:text-slate-400">
              Página {page} de {totalPages}
            </span>
            <button
              id="page_next"
              disabled={page >= totalPages}
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              className="tw:rounded-lg tw:border tw:border-black/15 tw:px-3 tw:py-1 tw:text-sm tw:disabled:cursor-not-allowed tw:disabled:opacity-40 tw:dark:border-white/15 tw:dark:text-slate-200"
            >
              Próxima
            </button>
          </div>
        </>
      )}
    </div>
  );
}
