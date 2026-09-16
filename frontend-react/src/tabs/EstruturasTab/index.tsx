import { useState } from "react";
import { useRegisterPrimaryAction } from "../../app/actionContext";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { pushToast } from "../../toast/toastStore";
import { Badge } from "../../ui/Badge";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { Field } from "../../ui/Field";
import { IconSitemap } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { TextInput } from "../../ui/TextInput";
import { computeContexto } from "./computeContexto";
import { useEstruturas } from "./useEstruturas";

function FileFeedback({ file, onClear }: { file: File | null; onClear: () => void }) {
  if (!file) return null;
  return (
    <div className="mt-3" aria-live="polite">
      <span className="inline-flex items-center gap-2 rounded-full border border-border bg-surface-2 py-1 pl-3 pr-1.5 text-sm text-text">
        {file.name}
        <button
          type="button"
          title="Remover arquivo"
          aria-label="Remover arquivo"
          onClick={(e) => {
            e.stopPropagation();
            onClear();
          }}
          className="flex h-5 w-5 items-center justify-center rounded-full text-text-subtle transition-colors hover:bg-danger/10 hover:text-danger"
        >
          ✕
        </button>
      </span>
    </div>
  );
}

export function EstruturasTab() {
  const e = useEstruturas();
  const isSubstituir = e.mode === "substituir";
  const allSelected = e.items.length > 0 && e.selectedIds.size === e.items.length;
  const [confirm, setConfirm] = useState<{ exportMode: "all" | "selected"; affected: { aprovacaoId: string }[] } | null>(
    null,
  );

  useRegisterPrimaryAction(e.previewLoading ? null : () => void e.preview());

  async function handleExport(exportMode: "all" | "selected") {
    const affected = await e.doExport(exportMode);
    if (affected) setConfirm({ exportMode, affected });
  }

  async function confirmExport() {
    if (!confirm) return;
    const { exportMode } = confirm;
    setConfirm(null);
    await e.doExport(exportMode, true);
  }

  return (
    <div>
      <PageHeader
        title="Estruturas de aprovação"
        description={
          isSubstituir
            ? "Substitua um aprovador por outro nas estruturas onde ele aparece, mantendo a posição/nível."
            : "Remova um aprovador específico das estruturas a partir do CPF e recompacte."
        }
        icon={<IconSitemap className="h-5 w-5" />}
      />

      <div className="mb-4 inline-flex rounded-lg border border-border bg-surface-2 p-1">
        <button
          type="button"
          id="aprovacao_mode_remover"
          onClick={() => e.setMode("remover")}
          className={`rounded-md px-3 py-1.5 text-sm font-medium transition-colors ${
            !isSubstituir ? "bg-surface text-text shadow-sm" : "text-text-muted hover:text-text"
          }`}
        >
          Remover
        </button>
        <button
          type="button"
          id="aprovacao_mode_substituir"
          onClick={() => e.setMode("substituir")}
          className={`rounded-md px-3 py-1.5 text-sm font-medium transition-colors ${
            isSubstituir ? "bg-surface text-text shadow-sm" : "text-text-muted hover:text-text"
          }`}
        >
          Substituir
        </button>
      </div>

      <Card>
        <div className="grid gap-4 sm:grid-cols-2">
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

        <div className="mt-4 flex flex-wrap items-end gap-4">
          <div className="min-w-[16rem]">
            <Field
              id="aprovacao_cpf"
              label={isSubstituir ? "CPF do aprovador atual" : "CPF do aprovador"}
              hint={<span id="aprovacao_cpf_help">Ex: 123.456.789-00 ou 12345678900.</span>}
            >
              <TextInput
                id="aprovacao_cpf"
                placeholder="Digite o CPF"
                aria-describedby="aprovacao_cpf_help"
                value={e.cpf}
                onChange={(ev) => e.setCpf(ev.target.value)}
              />
            </Field>
          </div>
          {isSubstituir && (
            <div className="min-w-[16rem]">
              <Field
                id="aprovacao_new_cpf"
                label="CPF do novo aprovador"
                hint={<span id="aprovacao_new_cpf_help">Quem assume o lugar do aprovador atual.</span>}
              >
                <TextInput
                  id="aprovacao_new_cpf"
                  placeholder="Digite o CPF"
                  aria-describedby="aprovacao_new_cpf_help"
                  value={e.newCpf}
                  onChange={(ev) => e.setNewCpf(ev.target.value)}
                />
              </Field>
            </div>
          )}
          <div className="flex items-center gap-3">
            <Button
              id="aprovacao_preview_btn"
              loading={e.previewLoading}
              onClick={() => e.preview()}
            >
              {e.previewLoading ? "Verificando..." : "Verificar"}
            </Button>
            {isSubstituir ? (
              <label className="flex items-center gap-1.5 text-sm text-text-muted">
                <input
                  type="checkbox"
                  id="aprovacao_replace_second_level"
                  checked={e.replaceSecondLevel}
                  onChange={(ev) => e.setReplaceSecondLevel(ev.target.checked)}
                />
                Substituir também no segundo nível
              </label>
            ) : (
              <label className="flex items-center gap-1.5 text-sm text-text-muted">
                <input
                  type="checkbox"
                  id="aprovacao_remove_second_level"
                  checked={e.removeSecondLevel}
                  onChange={(ev) => e.setRemoveSecondLevel(ev.target.checked)}
                />
                Remover também do segundo nível
              </label>
            )}
          </div>
          <div
            id="aprovacao_status"
            className={`text-sm ${e.status.isError ? "text-danger" : "text-text-muted"}`}
            aria-live="polite"
          >
            {e.status.message || e.readinessHint}
          </div>
        </div>

        {e.hasPreview && (
          <div id="aprovacao_preview_panel" className="mt-4 rounded-lg border border-border bg-surface-2 p-4">
            <div className="mb-3 flex flex-wrap items-start justify-between gap-2">
              <div className="flex flex-wrap gap-6">
                <div>
                  <p className="mb-0.5 text-xs uppercase tracking-wide text-text-subtle">
                    {isSubstituir ? "Aprovador atual" : "Aprovador localizado"}
                  </p>
                  <div id="aprovacao_approver_name" className="font-semibold text-text">
                    {e.approver?.nomeCompleto || "—"}
                  </div>
                  <div id="aprovacao_approver_cpf" className="text-sm text-text-muted">
                    {e.approver?.cpf ? `CPF: ${e.approver.cpf}` : ""}
                  </div>
                </div>
                {isSubstituir && (
                  <div>
                    <p className="mb-0.5 text-xs uppercase tracking-wide text-text-subtle">Novo aprovador</p>
                    <div id="aprovacao_new_approver_name" className="font-semibold text-text">
                      {e.newApprover?.nomeCompleto || "—"}
                    </div>
                    <div id="aprovacao_new_approver_cpf" className="text-sm text-text-muted">
                      {e.newApprover?.cpf ? `CPF: ${e.newApprover.cpf}` : ""}
                    </div>
                  </div>
                )}
              </div>
              <div className="text-right text-sm">
                <div>
                  <span className="text-text-muted">Estruturas afetadas: </span>
                  <span id="aprovacao_summary_estruturas" className="font-semibold text-text">
                    {e.summary?.estruturasAfetadas ?? 0}
                  </span>
                </div>
                <div>
                  <span className="text-text-muted">Ocorrências totais: </span>
                  <span id="aprovacao_summary_ocorrencias" className="font-semibold text-text">
                    {e.summary?.ocorrenciasTotal ?? 0}
                  </span>
                </div>
                <div id="aprovacao_summary_tipos" className="mt-1 flex flex-wrap justify-end gap-1">
                  {Object.entries(e.summary?.porAprovacaoPor || {}).map(([k, v]) => (
                    <Badge key={k}>
                      {k}: {v}
                    </Badge>
                  ))}
                </div>
              </div>
            </div>

            <div className="mb-2 flex flex-wrap items-center gap-2">
              {e.selectedIds.size > 0 && (
                <span id="aprovacao_selected_badge">
                  <Badge tone="neutral">
                    Selecionados: <span id="aprovacao_selected_count">{e.selectedIds.size}</span>
                  </Badge>
                </span>
              )}
              {isSubstituir ? (
                <>
                  <Button
                    id="aprovacao_substitute_all_btn"
                    variant="outline"
                    size="sm"
                    disabled={!e.items.length || e.exportLoading}
                    onClick={() => handleExport("all")}
                  >
                    Substituir em todas
                  </Button>
                  <Button
                    id="aprovacao_substitute_selected_btn"
                    variant="primary"
                    size="sm"
                    disabled={e.selectedIds.size === 0 || e.exportLoading}
                    onClick={() => handleExport("selected")}
                  >
                    Substituir selecionadas
                  </Button>
                </>
              ) : (
                <>
                  <Button
                    id="aprovacao_remove_all_btn"
                    variant="outline"
                    size="sm"
                    disabled={!e.items.length || e.exportLoading}
                    onClick={() => handleExport("all")}
                    className="border-danger text-danger hover:bg-danger/10"
                  >
                    Remover de todas
                  </Button>
                  <Button
                    id="aprovacao_remove_selected_btn"
                    variant="danger"
                    size="sm"
                    disabled={e.selectedIds.size === 0 || e.exportLoading}
                    onClick={() => handleExport("selected")}
                  >
                    Remover selecionadas
                  </Button>
                </>
              )}
            </div>

            {e.items.length > 0 && (
              <div id="aprovacao_table_wrap" className="overflow-x-auto rounded-lg border border-border">
                <table className="w-full text-sm">
                  <thead className="bg-surface-sunken">
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
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">AprovacaoId</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">AprovacaoPor</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Aprovacao</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Tipo</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Valor</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Descrição</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Posições</th>
                      <th scope="col" className="px-3 py-2 text-left font-medium text-text-muted">Segundo nível</th>
                    </tr>
                  </thead>
                  <tbody id="aprovacao_table_body">
                    {e.items.map((item) => {
                      const flagged = isSubstituir ? item.teraDuplicidade : item.ficaraSemAprovador;
                      const flagTitle = isSubstituir
                        ? "O novo aprovador já está presente nesta estrutura"
                        : "Esta estrutura ficará sem aprovadores";
                      return (
                        <tr
                          key={item.aprovacaoId}
                          className={`border-t border-border ${flagged ? "bg-warning/15" : ""}`}
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
                          <td className="px-3 py-1.5 text-text">
                            {item.aprovacaoId}
                            {flagged && (
                              <span title={flagTitle} className="ml-1 text-warning">
                                ⚠
                              </span>
                            )}
                          </td>
                          <td className="px-3 py-1.5 text-text">{item.aprovacaoPor}</td>
                          <td className="px-3 py-1.5 text-text">{item.aprovacao ?? ""}</td>
                          <td className="px-3 py-1.5 text-text">{item.tipo ?? ""}</td>
                          <td className="px-3 py-1.5 text-text">{item.valor ?? ""}</td>
                          <td className="px-3 py-1.5 text-text">{computeContexto(item)}</td>
                          <td className="px-3 py-1.5 text-text">
                            {item.posicoes?.length ? item.posicoes.join(", ") : "—"}
                          </td>
                          <td className="px-3 py-1.5 text-text">{item.segundoNivel ? "Sim" : "Não"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            <div className="mt-3 flex justify-end">
              <Button
                id="aprovacao_clear_all_btn"
                variant="danger"
                size="sm"
                title="Limpar todos os dados e resetar"
                onClick={e.clearAll}
              >
                Limpar Tudo
              </Button>
            </div>
          </div>
        )}
      </Card>

      <RunHistoryPanel operation="estruturas" />

      <Modal
        open={confirm !== null}
        onClose={() => {
          setConfirm(null);
          pushToast("Exportação cancelada.", "info");
        }}
        title="Atenção!"
        footer={
          <>
            <Button
              variant="outline"
              onClick={() => {
                setConfirm(null);
                pushToast("Exportação cancelada.", "info");
              }}
            >
              Cancelar
            </Button>
            <Button variant="danger" data-testid="aprovacao-confirm-continuar" onClick={confirmExport}>
              Sim, continuar
            </Button>
          </>
        }
      >
        <p>
          {isSubstituir
            ? "O novo aprovador já está presente em algumas estruturas:"
            : "Algumas estruturas ficarão sem nenhum aprovador após a remoção:"}
        </p>
        <ul className="my-2 max-h-52 list-disc overflow-y-auto pl-5">
          {confirm?.affected.map((s) => (
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
