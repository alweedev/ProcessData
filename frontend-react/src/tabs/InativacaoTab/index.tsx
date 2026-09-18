import { useMemo, useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { Field } from "../../ui/Field";
import { IconUserMinus } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { Skeleton } from "../../ui/Skeleton";
import { TextInput } from "../../ui/TextInput";
import { Textarea } from "../../ui/Textarea";
import { formatCpf, useInativacao } from "./useInativacao";

const PAGE_SIZE = 10;

function rowStatusClass(found: boolean, status: string): string {
  const s = (status || "").trim().toUpperCase();
  // Tokens `-soft` (não opacidade sobre a cor base) — a cor base muda de
  // tom entre claro/escuro (ex.: warning é marrom no claro, amarelo vivo no
  // escuro), então aplicar a mesma opacidade nela dava tintas de intensidade
  // bem diferente entre os temas. Os tokens `-soft` já são escolhidos por
  // tema pra ter o mesmo peso visual dos dois lados.
  if (!found) return "bg-warning-soft";
  if (s === "ATIVO") return "bg-success-soft";
  if (s) return "bg-danger-soft";
  return "";
}

export function InativacaoTab() {
  const inativacao = useInativacao();
  const [resultSearch, setResultSearch] = useState("");
  const [page, setPage] = useState(1);
  const [helpOpen, setHelpOpen] = useState(false);

  const filtered = useMemo(() => {
    const term = resultSearch.trim().toLowerCase();
    if (!term) return inativacao.results;
    return inativacao.results.filter((r) => r.nome.toLowerCase().includes(term) || r.cpf.includes(term));
  }, [inativacao.results, resultSearch]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const pageRows = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  function pickFile(file: File | null) {
    inativacao.setBaseFile(file);
    setPage(1);
  }

  function clearBase() {
    pickFile(null);
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
    <div>
      <PageHeader
        title="Inativação"
        description="Busque usuários numa base e gere o arquivo de inativação."
        icon={<IconUserMinus className="h-5 w-5" />}
        actions={
          <Button
            id="inativacao_help_btn"
            variant="outline"
            size="sm"
            aria-label="Como usar a inativação"
            title="Guia da inativação"
            onClick={() => setHelpOpen(true)}
          >
            Como usar
          </Button>
        }
      />

      <Modal open={helpOpen} onClose={() => setHelpOpen(false)} title="Inativação">
        <p className="mb-2 font-semibold text-text">Passos para inativação rápida</p>
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

      <Card>
        <FileDropzone
          id="inativacao_base"
          containerId="inativacao_base_uploadArea"
          accept=".xlsx,.xls"
          ariaLabel="Upload da base de inativação. Pressione para selecionar arquivo"
          description="Arraste a base (.xlsx) ou selecione o arquivo."
          syncFiles={inativacao.baseFile ? [inativacao.baseFile] : []}
          onFiles={(list) => pickFile(list[0] ?? null)}
        >
          <div id="inativacao_base_feedback" aria-live="polite" className="mt-3">
            {inativacao.baseFile && (
              <span className="inline-flex items-center gap-2 rounded-pill border border-border bg-surface-2 py-1 pl-3 pr-1.5 text-sm text-text">
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
                  className="flex h-5 w-5 items-center justify-center rounded-pill text-text-subtle transition-colors hover:bg-danger/10 hover:text-danger"
                >
                  ✕
                </button>
              </span>
            )}
          </div>
        </FileDropzone>

        <div className="mt-4">
          <Field id="lista_text" label="Informe abaixo como deseja localizar os usuários.">
            <Textarea
              id="lista_text"
              rows={4}
              placeholder={"Ex: João Silva\n12345678901\nusuario@example.com"}
              aria-describedby="lista_valid_summary lista_duplicates_warning"
              title="Insira (nomes, CPFs ou e-mails) e eles serão validados automaticamente"
              value={inativacao.listText}
              onChange={(e) => inativacao.setListText(e.target.value)}
            />
          </Field>
          <div id="lista_valid_summary" className="mt-2 text-sm text-text-muted">
            <span className="rounded bg-accent/10 px-2 py-0.5 font-medium text-accent">
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
          <div id="inativacao_progress" className="mt-3 h-2 overflow-hidden rounded-pill bg-surface-sunken">
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

        <div className="mt-4">
          <Button
            id="inativacao_btn"
            aria-label="Buscar usuários para inativação"
            title="Executar a busca de itens digitados na base"
            disabled={!inativacao.canSearch}
            loading={inativacao.searching}
            onClick={handleSearch}
          >
            {inativacao.searching ? "Buscando..." : "Buscar Usuários"}
          </Button>
        </div>

        <div id="inativacao_status" className="mt-3 text-sm text-text-muted" aria-live="polite" />
        {inativacao.debugMsg && (
          <div id="inativacao_debug" className="mt-3 text-sm text-danger" aria-live="assertive">
            {inativacao.debugMsg}
          </div>
        )}

        {showResults && (
          <>
            <div id="inativacao_results_controls" className="mt-4 flex flex-wrap items-center gap-2">
              <TextInput
                id="result_search"
                placeholder="Filtrar por nome ou CPF"
                aria-label="Filtrar resultados"
                className="min-w-[220px] flex-1"
                value={resultSearch}
                onChange={(e) => {
                  setResultSearch(e.target.value);
                  setPage(1);
                }}
              />
              {inativacao.results.length > 0 && (
                <Button
                  id="inativacao_generate_btn"
                  variant="success"
                  aria-label="Gerar inativação"
                  title="Processar e baixar o relatório final"
                  loading={inativacao.generating}
                  onClick={handleGenerate}
                >
                  {inativacao.generating ? "Gerando..." : "Gerar Inativação"}
                </Button>
              )}
            </div>

            <div id="inativacao_results" className="mt-3 overflow-x-auto rounded-surface border border-border">
              <table id="results_table" className="w-full text-sm">
                <thead className="bg-surface-sunken">
                  <tr>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Nome Completo</th>
                    <th scope="col" className="whitespace-nowrap px-3 py-2 text-left font-medium text-text-muted">CPF</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">E-mail</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Status Atual</th>
                  </tr>
                </thead>
                <tbody id="results_body">
                  {inativacao.searching && pageRows.length === 0 && (
                    <tr aria-hidden="true">
                      <td colSpan={4} className="px-3 py-2">
                        <div className="space-y-1.5">
                          <Skeleton className="h-5 w-full" />
                          <Skeleton className="h-5 w-full" />
                          <Skeleton className="h-5 w-full" />
                        </div>
                      </td>
                    </tr>
                  )}
                  {pageRows.map((r, i) => (
                    <tr
                      key={`${r.cpf}-${i}`}
                      data-cpf={r.cpf}
                      className={`border-t border-border ${rowStatusClass(r.found, r.status_atual)}`}
                    >
                      <td className="px-3 py-1.5 text-text">{r.nome}</td>
                      <td className="px-3 py-1.5 font-mono text-text">{formatCpf(r.cpf)}</td>
                      <td className="px-3 py-1.5 text-text">{r.email}</td>
                      <td className="px-3 py-1.5 text-text">{r.status_atual}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div id="results_pagination" className="mt-2 flex items-center justify-between">
              <Button
                id="page_prev"
                variant="outline"
                size="sm"
                disabled={page <= 1}
                onClick={() => setPage((p) => Math.max(1, p - 1))}
              >
                Anterior
              </Button>
              <span id="page_info" className="text-sm text-text-muted">
                Página {page} de {totalPages}
              </span>
              <Button
                id="page_next"
                variant="outline"
                size="sm"
                disabled={page >= totalPages}
                onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              >
                Próxima
              </Button>
            </div>
          </>
        )}
      </Card>

      <RunHistoryPanel operation="inativacao" />
    </div>
  );
}
