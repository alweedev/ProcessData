import { useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { useRegisterPrimaryAction } from "../../app/actionContext";
import { usePersistedState } from "../../hooks/usePersistedState";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { Field } from "../../ui/Field";
import { FileChips } from "../../ui/FileChips";
import { IconUpload } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { Select } from "../../ui/Select";
import { ValidationReport } from "../../ui/ValidationReport";
import { useCadastro } from "./useCadastro";

export function CadastroTab() {
  const cadastro = useCadastro();
  const [prefs, setPrefs] = usePersistedState("cadastro_prefs", { login_choice: "CPF", fluxo: "SELF" });
  const [resetKey, setResetKey] = useState(0);
  const [helpOpen, setHelpOpen] = useState(false);

  async function handleSubmit() {
    await cadastro.submit(prefs.login_choice, prefs.fluxo, () => setResetKey((k) => k + 1));
  }

  useRegisterPrimaryAction(cadastro.generating ? null : () => void handleSubmit());

  function clear() {
    cadastro.clear();
    setResetKey((k) => k + 1);
  }

  return (
    <div>
      <PageHeader
        title="Cadastro em massa"
        description="Trate a planilha de fichas e gere o arquivo pronto para carga."
        icon={<IconUpload className="h-5 w-5" />}
        actions={
          <Button
            id="cadastro_help_btn"
            variant="outline"
            size="sm"
            aria-label="Como usar o cadastro em lote"
            title="Guia do cadastro em lote"
            onClick={() => setHelpOpen(true)}
          >
            Como usar
          </Button>
        }
      />

      <Modal open={helpOpen} onClose={() => setHelpOpen(false)} title="Como usar carga cadastro">
        <ol className="list-decimal space-y-1 pl-5">
          <li>
            <strong>1º Passo:</strong> carregue a planilha Excel (<strong>.xlsx/.xls</strong>) já preenchida.
          </li>
          <li>
            <strong>2º Passo:</strong> escolha o <strong>tipo de login</strong> (CPF ou e-mail) e o <strong>fluxo</strong> (SELF ou FRONT).
          </li>
          <li>
            <strong>3º Passo:</strong> clique em <em>Gerar</em> para processar e obter o arquivo pronto para carga.
          </li>
          <li>
            <strong>4º Passo:</strong> acompanhe e recupere execuções no <strong>Histórico</strong> quando precisar.
          </li>
        </ol>
        <p className="mt-2 text-text-subtle">Dica: valide a planilha antes de gerar para evitar retrabalho.</p>
      </Modal>

      <Card>
        <FileDropzone
          key={resetKey}
          id="cadastro_files"
          containerId="cadastro_uploadArea"
          accept=".xlsx,.xls"
          multiple
          ariaLabel="Upload da base de cadastro. Pressione para selecionar arquivo"
          description="Arraste o Excel (.xlsx) ou pressione para selecionar."
          syncFiles={cadastro.files}
          onFiles={(list) => {
            const accepted = cadastro.pickFiles(list);
            if (!accepted) setResetKey((k) => k + 1);
          }}
        >
          <div id="cadastro_uploadFeedback" aria-live="polite">
            <FileChips
              files={cadastro.files}
              onRemove={cadastro.removeFile}
              onClearAll={clear}
              clearAllId="cadastro_clear_btn"
            />
          </div>
        </FileDropzone>

        <div className="mt-4 grid gap-3 sm:grid-cols-2">
          <Field id="cadastro_login_choice" label="Tipo de login">
            <Select
              id="cadastro_login_choice"
              aria-label="Escolher tipo de login"
              title="Selecionar formato principal de acesso"
              value={prefs.login_choice}
              onChange={(e) => setPrefs({ login_choice: e.target.value })}
            >
              <option value="CPF">CPF</option>
              <option value="EMAIL">E-MAIL</option>
            </Select>
          </Field>
          <Field id="cadastro_fluxo" label="Fluxo">
            <Select
              id="cadastro_fluxo"
              aria-label="Escolher fluxo"
              title="Definir processo operacional (SELF ou FRONT)"
              value={prefs.fluxo}
              onChange={(e) => setPrefs({ fluxo: e.target.value })}
            >
              <option value="SELF">SELF</option>
              <option value="FRONT">FRONT</option>
            </Select>
          </Field>
        </div>

        <div className="mt-4 space-y-3">
          <Button
            id="cadastro_validate_btn"
            variant="secondary"
            size="sm"
            disabled={cadastro.files.length === 0}
            loading={cadastro.validation.status === "loading"}
            onClick={() => cadastro.validate(prefs.login_choice, prefs.fluxo)}
          >
            {cadastro.validation.status === "loading" ? "Validando..." : "Validar planilha"}
          </Button>
          {cadastro.validation.status !== "idle" && (
            <ValidationReport
              report={cadastro.validation.report}
              loading={cadastro.validation.status === "loading"}
              error={cadastro.validation.status === "error" ? cadastro.validation.error : null}
            />
          )}
        </div>

        {cadastro.generating && (
          <div id="cadastro_progress" className="mt-4 h-2 overflow-hidden rounded-full bg-surface-sunken">
            <div
              id="cadastro_progressBar"
              role="progressbar"
              aria-valuenow={Math.round(cadastro.progress)}
              aria-valuemin={0}
              aria-valuemax={100}
              style={{ width: `${cadastro.progress}%` }}
              className="h-full bg-accent transition-[width]"
            />
          </div>
        )}

        <div className="mt-4">
          <Button
            id="cadastro_btn"
            aria-label="Gerar cadastro"
            title="Processar a planilha e gerar arquivo tratado"
            loading={cadastro.generating}
            onClick={handleSubmit}
          >
            {cadastro.generating ? "Processando..." : "Gerar"}
          </Button>
        </div>

        <div id="cadastro_status" className="mt-3 text-sm" aria-live="polite">
          {cadastro.done && <span className="font-medium text-success">✔ Concluído</span>}
        </div>
        {cadastro.debugMsg && (
          <div id="cadastro_debug" className="mt-3 text-sm text-danger" aria-live="assertive">
            {cadastro.debugMsg}
          </div>
        )}
      </Card>

      <RunHistoryPanel operation="cadastro" />
    </div>
  );
}
