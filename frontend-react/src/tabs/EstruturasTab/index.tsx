import { useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { pushToast } from "../../toast/toastStore";
import { useEstruturas, type PreviewItem } from "./useEstruturas";

function computeContexto(item: PreviewItem): string {
  const cod = item.ccCodigo || "";
  const desc = item.ccDescricao || "";
  if (item.aprovacaoPor === "VIAJANTE" && item.viajanteNomeCompleto) return item.viajanteNomeCompleto;
  if (item.aprovacaoPor === "CCEMPRESA") {
    if (cod && desc) return `${cod} - ${desc}`;
    return cod || desc || "";
  }
  if (cod && desc) return `${cod} - ${desc}`;
  if (item.viajanteNomeCompleto) return item.viajanteNomeCompleto;
  return cod || desc || "";
}

function FileFeedback({ file, onClear }: { file: File | null; onClear: () => void }) {
  if (!file) return null;
  return (
    <div className="mt-3" aria-live="polite">
      <span className="inline-flex items-center gap-2 rounded-full bg-success/10 px-3 py-1 text-sm text-success">
        {file.name}
        <button
          type="button"
          title="Remover arquivo"
          aria-label="Remover arquivo"
          onClick={(e) => {
            e.stopPropagation();
            onClear();
          }}
          className="text-success hover:text-danger"
        >
          ✕
        </button>
      </span>
    </div>
  );
}

export function EstruturasTab() {
  const e = useEstruturas();
  const allSelected = e.items.length > 0 && e.selectedIds.size === e.items.length;
  const [confirm, setConfirm] = useState<{ mode: "all" | "selected"; empty: { aprovacaoId: string }[] } | null>(null);

  async function handleExport(mode: "all" | "selected") {
    const emptyStructures = await e.doExport(mode);
    if (emptyStructures) setConfirm({ mode, empty: emptyStructures });
  }

  async function confirmExport() {
    if (!confirm) return;
    const { mode } = confirm;
    setConfirm(null);
    await e.doExport(mode, true);
  }

  return (
    <div className="rounded-xl border border-black/10 bg-surface p-4 shadow-sm dark:border-white/10 dark:bg-surface-dark sm:p-6">
      <h2 className="mb-1 text-lg font-semibold text-slate-900 dark:text-slate-100">Estruturas de aprovação</h2>
      <p className="mb-4 text-sm text-slate-600 dark:text-slate-300">
        Remova um aprovador específico de estruturas de aprovação a partir do CPF.
      </p>

      <div className="mb-4 grid gap-4 sm:grid-cols-2">
        <FileDropzone
          id="aprovacao_users_file"
          containerId="aprovacao_users_uploadArea"
          accept=".xlsx,.xls"
          ariaLabel="Upload da base de usuários. Pressione para selecionar arquivo"
          description="Base de Usuários (.xlsx/.xls) — arraste ou pressione para selecionar."
          onFiles={(list) => e.setUsersFile(list[0] ?? null)}
        >
          <div id="aprovacao_users_feedback">
            <FileFeedback file={e.usersFile} onClear={() => e.setUsersFile(null)} />
          </div>
        </FileDropzone>
        <FileDropzone
          id="aprovacao_base_file"
          containerId="aprovacao_base_uploadArea"
          accept=".xlsx,.xls"
          ariaLabel="Upload da base de aprovação. Pressione para selecionar arquivo"
          description="Base de Carga Aprovação (.xlsx/.xls) — arraste ou pressione para selecionar."
          onFiles={(list) => e.setBaseFile(list[0] ?? null)}
        >
          <div id="aprovacao_base_feedback">
            <FileFeedback file={e.baseFile} onClear={() => e.setBaseFile(null)} />
          </div>
        </FileDropzone>
      </div>

      <div className="mb-4 flex flex-wrap items-end gap-4">
        <div>
          <label htmlFor="aprovacao_cpf" className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-200">
            CPF do aprovador
          </label>
          <input
            type="text"
            id="aprovacao_cpf"
            placeholder="Digite o CPF"
            aria-describedby="aprovacao_cpf_help"
            value={e.cpf}
            onChange={(ev) => e.setCpf(ev.target.value)}
            className="rounded-lg border border-black/10 bg-white px-3 py-1.5 text-sm text-slate-900 outline-none focus:border-accent dark:border-white/10 dark:bg-surface-dark-alt dark:text-slate-100"
          />
          <div id="aprovacao_cpf_help" className="mt-1 text-xs text-slate-500 dark:text-slate-400">
            Ex: 123.456.789-00 ou 12345678900.
          </div>
        </div>
        <div className="flex items-center gap-3">
          <button
            type="button"
            id="aprovacao_preview_btn"
            disabled={e.previewLoading}
            onClick={() => e.preview()}
            className="rounded-lg bg-accent px-4 py-2 text-sm font-medium text-white hover:bg-accent-hover disabled:cursor-not-allowed disabled:opacity-50"
          >
            {e.previewLoading ? "Verificando..." : "Verificar"}
          </button>
          <label className="flex items-center gap-1.5 text-sm text-slate-600 dark:text-slate-300">
            <input
              type="checkbox"
              id="aprovacao_remove_second_level"
              checked={e.removeSecondLevel}
              onChange={(ev) => e.setRemoveSecondLevel(ev.target.checked)}
            />
            Remover também do segundo nível
          </label>
        </div>
        <div id="aprovacao_status" className={`text-sm ${e.status.isError ? "text-danger" : "text-slate-500 dark:text-slate-400"}`} aria-live="polite">
          {e.status.message || e.readinessHint}
        </div>
      </div>

      {e.hasPreview && (
        <div id="aprovacao_preview_panel" className="rounded-lg border border-black/10 p-4 dark:border-white/10">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <div>
              <p className="mb-0.5 text-xs uppercase text-slate-500 dark:text-slate-400">Aprovador localizado</p>
              <div id="aprovacao_approver_name" className="font-semibold text-slate-900 dark:text-slate-100">
                {e.approver?.nomeCompleto || "—"}
              </div>
              <div id="aprovacao_approver_cpf" className="text-sm text-slate-500 dark:text-slate-400">
                {e.approver?.cpf ? `CPF: ${e.approver.cpf}` : ""}
              </div>
            </div>
            <div className="text-right text-sm">
              <div>
                <span className="text-slate-500 dark:text-slate-400">Estruturas afetadas: </span>
                <span id="aprovacao_summary_estruturas" className="font-semibold text-slate-900 dark:text-slate-100">
                  {e.summary?.estruturasAfetadas ?? 0}
                </span>
              </div>
              <div>
                <span className="text-slate-500 dark:text-slate-400">Ocorrências totais: </span>
                <span id="aprovacao_summary_ocorrencias" className="font-semibold text-slate-900 dark:text-slate-100">
                  {e.summary?.ocorrenciasTotal ?? 0}
                </span>
              </div>
              <div id="aprovacao_summary_tipos" className="mt-1 flex justify-end gap-1">
                {Object.entries(e.summary?.porAprovacaoPor || {}).map(([k, v]) => (
                  <span key={k} className="rounded border border-black/10 bg-black/[.03] px-1.5 py-0.5 text-xs dark:border-white/10 dark:bg-white/[.04]">
                    {k}: {v}
                  </span>
                ))}
              </div>
            </div>
          </div>

          <div className="mb-2 flex flex-wrap items-center gap-2">
            {e.selectedIds.size > 0 && (
              <span id="aprovacao_selected_badge" className="rounded bg-black/10 px-2 py-0.5 text-xs dark:bg-white/10">
                Selecionados: <span id="aprovacao_selected_count">{e.selectedIds.size}</span>
              </span>
            )}
            <button
              type="button"
              id="aprovacao_remove_all_btn"
              disabled={!e.items.length || e.exportLoading}
              onClick={() => handleExport("all")}
              className="rounded-lg border border-danger px-3 py-1.5 text-sm font-medium text-danger hover:bg-danger/10 disabled:cursor-not-allowed disabled:opacity-40"
            >
              Remover de todas
            </button>
            <button
              type="button"
              id="aprovacao_remove_selected_btn"
              disabled={e.selectedIds.size === 0 || e.exportLoading}
              onClick={() => handleExport("selected")}
              className="rounded-lg bg-danger px-3 py-1.5 text-sm font-medium text-white hover:brightness-95 disabled:cursor-not-allowed disabled:opacity-40"
            >
              Remover selecionadas
            </button>
          </div>

          {e.items.length > 0 && (
            <div id="aprovacao_table_wrap" className="overflow-x-auto rounded-lg border border-black/10 dark:border-white/10">
              <table className="w-full text-sm">
                <thead className="bg-black/[.03] dark:bg-white/[.04]">
                  <tr>
                    <th scope="col" className="w-8 px-2 py-2">
                      <input
                        type="checkbox"
                        id="aprovacao_check_all"
                        aria-label="Selecionar todas as estruturas"
                        checked={allSelected}
                        onChange={(ev) => e.toggleSelectAll(ev.target.checked)}
                      />
                    </th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">AprovacaoId</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">AprovacaoPor</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">Aprovacao</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">Tipo</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">Valor</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">Descrição</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">Posições</th>
                    <th scope="col" className="px-3 py-2 text-left font-medium text-slate-600 dark:text-slate-300">Segundo nível</th>
                  </tr>
                </thead>
                <tbody id="aprovacao_table_body">
                  {e.items.map((item) => (
                    <tr
                      key={item.aprovacaoId}
                      className={`border-t border-black/5 dark:border-white/5 ${item.ficaraSemAprovador ? "bg-warning/15" : ""}`}
                    >
                      <td className="px-2 py-1.5">
                        <input
                          type="checkbox"
                          className="aprov-row-check"
                          data-id={item.aprovacaoId}
                          checked={e.selectedIds.has(item.aprovacaoId)}
                          onChange={(ev) => e.toggleSelected(item.aprovacaoId, ev.target.checked)}
                        />
                      </td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">
                        {item.aprovacaoId}
                        {item.ficaraSemAprovador && (
                          <span title="Esta estrutura ficará sem aprovadores" className="ml-1 text-warning">
                            ⚠
                          </span>
                        )}
                      </td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{item.aprovacaoPor}</td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{item.aprovacao ?? ""}</td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{item.tipo ?? ""}</td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{item.valor ?? ""}</td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{computeContexto(item)}</td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">
                        {item.posicoes?.length ? item.posicoes.join(", ") : "—"}
                      </td>
                      <td className="px-3 py-1.5 text-slate-800 dark:text-slate-100">{item.segundoNivel ? "Sim" : "Não"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          <div className="mt-3 flex justify-end">
            <button
              type="button"
              id="aprovacao_clear_all_btn"
              title="Limpar todos os dados e resetar"
              onClick={e.clearAll}
              className="flex items-center gap-1.5 rounded-lg bg-gradient-to-r from-[#ff416c] to-[#ff4b2b] px-3 py-1.5 text-sm font-medium text-white hover:brightness-105"
            >
              Limpar Tudo
            </button>
          </div>
        </div>
      )}

      <Modal
        open={confirm !== null}
        onClose={() => {
          setConfirm(null);
          pushToast("Exportação cancelada.", "info");
        }}
        title="Atenção!"
        footer={
          <>
            <button
              type="button"
              onClick={() => {
                setConfirm(null);
                pushToast("Exportação cancelada.", "info");
              }}
              className="rounded-lg border border-black/15 px-4 py-1.5 text-sm font-medium dark:border-white/15 dark:text-slate-200"
            >
              Cancelar
            </button>
            <button
              type="button"
              data-testid="aprovacao-confirm-continuar"
              onClick={confirmExport}
              className="rounded-lg bg-danger px-4 py-1.5 text-sm font-medium text-white hover:brightness-95"
            >
              Sim, continuar
            </button>
          </>
        }
      >
        <p>Algumas estruturas ficarão sem nenhum aprovador após a remoção:</p>
        <ul className="my-2 max-h-52 list-disc overflow-y-auto pl-5">
          {confirm?.empty.map((s) => (
            <li key={s.aprovacaoId}>
              AprovacaoId: <strong>{s.aprovacaoId}</strong>
            </li>
          ))}
        </ul>
        <p className="font-bold text-danger">Deseja continuar mesmo assim?</p>
      </Modal>
    </div>
  );
}
