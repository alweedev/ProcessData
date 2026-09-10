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
    <div className="tw:mt-3" aria-live="polite">
      <span className="tw:inline-flex tw:items-center tw:gap-2 tw:rounded-full tw:bg-success/10 tw:px-3 tw:py-1 tw:text-sm tw:text-success">
        {file.name}
        <button
          type="button"
          title="Remover arquivo"
          aria-label="Remover arquivo"
          onClick={(e) => {
            e.stopPropagation();
            onClear();
          }}
          className="tw:text-success tw:hover:text-danger"
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
    <div className="tw:rounded-xl tw:border tw:border-black/10 tw:bg-surface tw:p-4 tw:shadow-sm tw:dark:border-white/10 tw:dark:bg-surface-dark tw:sm:p-6">
      <h2 className="tw:mb-1 tw:text-lg tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">Estruturas de aprovação</h2>
      <p className="tw:mb-4 tw:text-sm tw:text-slate-600 tw:dark:text-slate-300">
        Remova um aprovador específico de estruturas de aprovação a partir do CPF.
      </p>

      <div className="tw:mb-4 tw:grid tw:gap-4 tw:sm:grid-cols-2">
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

      <div className="tw:mb-4 tw:flex tw:flex-wrap tw:items-end tw:gap-4">
        <div>
          <label htmlFor="aprovacao_cpf" className="tw:mb-1 tw:block tw:text-sm tw:font-medium tw:text-slate-700 tw:dark:text-slate-200">
            CPF do aprovador
          </label>
          <input
            type="text"
            id="aprovacao_cpf"
            placeholder="Digite o CPF"
            aria-describedby="aprovacao_cpf_help"
            value={e.cpf}
            onChange={(ev) => e.setCpf(ev.target.value)}
            className="tw:rounded-lg tw:border tw:border-black/10 tw:bg-white tw:px-3 tw:py-1.5 tw:text-sm tw:text-slate-900 tw:outline-none tw:focus:border-accent tw:dark:border-white/10 tw:dark:bg-surface-dark-alt tw:dark:text-slate-100"
          />
          <div id="aprovacao_cpf_help" className="tw:mt-1 tw:text-xs tw:text-slate-500 tw:dark:text-slate-400">
            Ex: 123.456.789-00 ou 12345678900.
          </div>
        </div>
        <div className="tw:flex tw:items-center tw:gap-3">
          <button
            type="button"
            id="aprovacao_preview_btn"
            disabled={e.previewLoading}
            onClick={() => e.preview()}
            className="tw:rounded-lg tw:bg-accent tw:px-4 tw:py-2 tw:text-sm tw:font-medium tw:text-white tw:hover:bg-accent-hover tw:disabled:cursor-not-allowed tw:disabled:opacity-50"
          >
            {e.previewLoading ? "Verificando..." : "Verificar"}
          </button>
          <label className="tw:flex tw:items-center tw:gap-1.5 tw:text-sm tw:text-slate-600 tw:dark:text-slate-300">
            <input
              type="checkbox"
              id="aprovacao_remove_second_level"
              checked={e.removeSecondLevel}
              onChange={(ev) => e.setRemoveSecondLevel(ev.target.checked)}
            />
            Remover também do segundo nível
          </label>
        </div>
        <div id="aprovacao_status" className={`tw:text-sm ${e.status.isError ? "tw:text-danger" : "tw:text-slate-500 tw:dark:text-slate-400"}`} aria-live="polite">
          {e.status.message || e.readinessHint}
        </div>
      </div>

      {e.hasPreview && (
        <div id="aprovacao_preview_panel" className="tw:rounded-lg tw:border tw:border-black/10 tw:p-4 tw:dark:border-white/10">
          <div className="tw:mb-3 tw:flex tw:flex-wrap tw:items-center tw:justify-between tw:gap-2">
            <div>
              <p className="tw:mb-0.5 tw:text-xs tw:uppercase tw:text-slate-500 tw:dark:text-slate-400">Aprovador localizado</p>
              <div id="aprovacao_approver_name" className="tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">
                {e.approver?.nomeCompleto || "—"}
              </div>
              <div id="aprovacao_approver_cpf" className="tw:text-sm tw:text-slate-500 tw:dark:text-slate-400">
                {e.approver?.cpf ? `CPF: ${e.approver.cpf}` : ""}
              </div>
            </div>
            <div className="tw:text-right tw:text-sm">
              <div>
                <span className="tw:text-slate-500 tw:dark:text-slate-400">Estruturas afetadas: </span>
                <span id="aprovacao_summary_estruturas" className="tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">
                  {e.summary?.estruturasAfetadas ?? 0}
                </span>
              </div>
              <div>
                <span className="tw:text-slate-500 tw:dark:text-slate-400">Ocorrências totais: </span>
                <span id="aprovacao_summary_ocorrencias" className="tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">
                  {e.summary?.ocorrenciasTotal ?? 0}
                </span>
              </div>
              <div id="aprovacao_summary_tipos" className="tw:mt-1 tw:flex tw:justify-end tw:gap-1">
                {Object.entries(e.summary?.porAprovacaoPor || {}).map(([k, v]) => (
                  <span key={k} className="tw:rounded tw:border tw:border-black/10 tw:bg-black/[.03] tw:px-1.5 tw:py-0.5 tw:text-xs tw:dark:border-white/10 tw:dark:bg-white/[.04]">
                    {k}: {v}
                  </span>
                ))}
              </div>
            </div>
          </div>

          <div className="tw:mb-2 tw:flex tw:flex-wrap tw:items-center tw:gap-2">
            {e.selectedIds.size > 0 && (
              <span id="aprovacao_selected_badge" className="tw:rounded tw:bg-black/10 tw:px-2 tw:py-0.5 tw:text-xs tw:dark:bg-white/10">
                Selecionados: <span id="aprovacao_selected_count">{e.selectedIds.size}</span>
              </span>
            )}
            <button
              type="button"
              id="aprovacao_remove_all_btn"
              disabled={!e.items.length || e.exportLoading}
              onClick={() => handleExport("all")}
              className="tw:rounded-lg tw:border tw:border-danger tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-danger tw:hover:bg-danger/10 tw:disabled:cursor-not-allowed tw:disabled:opacity-40"
            >
              Remover de todas
            </button>
            <button
              type="button"
              id="aprovacao_remove_selected_btn"
              disabled={e.selectedIds.size === 0 || e.exportLoading}
              onClick={() => handleExport("selected")}
              className="tw:rounded-lg tw:bg-danger tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-white tw:hover:brightness-95 tw:disabled:cursor-not-allowed tw:disabled:opacity-40"
            >
              Remover selecionadas
            </button>
          </div>

          {e.items.length > 0 && (
            <div id="aprovacao_table_wrap" className="tw:overflow-x-auto tw:rounded-lg tw:border tw:border-black/10 tw:dark:border-white/10">
              <table className="tw:w-full tw:text-sm">
                <thead className="tw:bg-black/[.03] tw:dark:bg-white/[.04]">
                  <tr>
                    <th scope="col" className="tw:w-8 tw:px-2 tw:py-2">
                      <input
                        type="checkbox"
                        id="aprovacao_check_all"
                        aria-label="Selecionar todas as estruturas"
                        checked={allSelected}
                        onChange={(ev) => e.toggleSelectAll(ev.target.checked)}
                      />
                    </th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">AprovacaoId</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">AprovacaoPor</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">Aprovacao</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">Tipo</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">Valor</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">Descrição</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">Posições</th>
                    <th scope="col" className="tw:px-3 tw:py-2 tw:text-left tw:font-medium tw:text-slate-600 tw:dark:text-slate-300">Segundo nível</th>
                  </tr>
                </thead>
                <tbody id="aprovacao_table_body">
                  {e.items.map((item) => (
                    <tr
                      key={item.aprovacaoId}
                      className={`tw:border-t tw:border-black/5 tw:dark:border-white/5 ${item.ficaraSemAprovador ? "tw:bg-warning/15" : ""}`}
                    >
                      <td className="tw:px-2 tw:py-1.5">
                        <input
                          type="checkbox"
                          className="aprov-row-check"
                          data-id={item.aprovacaoId}
                          checked={e.selectedIds.has(item.aprovacaoId)}
                          onChange={(ev) => e.toggleSelected(item.aprovacaoId, ev.target.checked)}
                        />
                      </td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">
                        {item.aprovacaoId}
                        {item.ficaraSemAprovador && (
                          <span title="Esta estrutura ficará sem aprovadores" className="tw:ml-1 tw:text-warning">
                            ⚠
                          </span>
                        )}
                      </td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{item.aprovacaoPor}</td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{item.aprovacao ?? ""}</td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{item.tipo ?? ""}</td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{item.valor ?? ""}</td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{computeContexto(item)}</td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">
                        {item.posicoes?.length ? item.posicoes.join(", ") : "—"}
                      </td>
                      <td className="tw:px-3 tw:py-1.5 tw:text-slate-800 tw:dark:text-slate-100">{item.segundoNivel ? "Sim" : "Não"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          <div className="tw:mt-3 tw:flex tw:justify-end">
            <button
              type="button"
              id="aprovacao_clear_all_btn"
              title="Limpar todos os dados e resetar"
              onClick={e.clearAll}
              className="tw:flex tw:items-center tw:gap-1.5 tw:rounded-lg tw:bg-gradient-to-r tw:from-[#ff416c] tw:to-[#ff4b2b] tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-white tw:hover:brightness-105"
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
              className="tw:rounded-lg tw:border tw:border-black/15 tw:px-4 tw:py-1.5 tw:text-sm tw:font-medium tw:dark:border-white/15 tw:dark:text-slate-200"
            >
              Cancelar
            </button>
            <button
              type="button"
              data-testid="aprovacao-confirm-continuar"
              onClick={confirmExport}
              className="tw:rounded-lg tw:bg-danger tw:px-4 tw:py-1.5 tw:text-sm tw:font-medium tw:text-white tw:hover:brightness-95"
            >
              Sim, continuar
            </button>
          </>
        }
      >
        <p>Algumas estruturas ficarão sem nenhum aprovador após a remoção:</p>
        <ul className="tw:my-2 tw:max-h-52 tw:list-disc tw:overflow-y-auto tw:pl-5">
          {confirm?.empty.map((s) => (
            <li key={s.aprovacaoId}>
              AprovacaoId: <strong>{s.aprovacaoId}</strong>
            </li>
          ))}
        </ul>
        <p className="tw:font-bold tw:text-danger">Deseja continuar mesmo assim?</p>
      </Modal>
    </div>
  );
}
