import { useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { usePersistedState } from "../../hooks/usePersistedState";
import { useCadastro } from "./useCadastro";

export function CadastroTab() {
  const cadastro = useCadastro();
  const [prefs, setPrefs] = usePersistedState("cadastro_prefs", { login_choice: "CPF", fluxo: "SELF" });
  const [resetKey, setResetKey] = useState(0);
  const [helpOpen, setHelpOpen] = useState(false);

  function clear() {
    cadastro.clear();
    setResetKey((k) => k + 1);
  }

  async function handleSubmit() {
    await cadastro.submit(prefs.login_choice, prefs.fluxo, () => setResetKey((k) => k + 1));
  }

  return (
    <div className="rounded-xl border border-black/10 bg-surface p-4 shadow-sm dark:border-white/10 dark:bg-surface-dark sm:p-6">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <h2 className="text-lg font-semibold text-slate-900 dark:text-slate-100">Cadastro Carga</h2>
        <button
          id="cadastro_help_btn"
          type="button"
          aria-label="Como usar o cadastro em lote"
          title="Guia do cadastro em lote"
          onClick={() => setHelpOpen(true)}
          className="rounded-lg border border-info px-3 py-1.5 text-sm font-medium text-info hover:bg-info/10"
        >
          Como usar
        </button>
      </div>

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
        <p className="mt-2 text-slate-500 dark:text-slate-400">
          Dica: revise os dados antes de enviar para evitar retrabalho.
        </p>
      </Modal>

      <FileDropzone
        key={resetKey}
        id="cadastro_files"
        containerId="cadastro_uploadArea"
        accept=".xlsx,.xls"
        multiple
        ariaLabel="Upload da base de cadastro. Pressione para selecionar arquivo"
        description="Arraste o Excel (.xlsx) ou pressione para selecionar."
        onFiles={(list) => {
          const accepted = cadastro.pickFiles(list);
          if (!accepted) setResetKey((k) => k + 1);
        }}
      >
        <div id="cadastro_uploadFeedback" aria-live="polite" className="mt-3">
          {cadastro.files.length > 0 && (
            <span className="inline-flex items-center gap-2 rounded-full bg-success/10 px-3 py-1 text-sm text-success">
              {cadastro.files.map((f) => f.name).join(", ")}
              <button
                id="cadastro_clear_btn"
                type="button"
                title="Remover arquivos"
                aria-label="Limpar seleção de arquivos"
                onClick={(e) => {
                  e.stopPropagation();
                  clear();
                }}
                className="text-success hover:text-danger"
              >
                ✕
              </button>
            </span>
          )}
        </div>
      </FileDropzone>

      <div className="mt-4 grid gap-3 sm:grid-cols-2">
        <div>
          <label htmlFor="cadastro_login_choice" className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-200">
            Tipo de login
          </label>
          <select
            id="cadastro_login_choice"
            aria-label="Escolher tipo de login"
            title="Selecionar formato principal de acesso"
            value={prefs.login_choice}
            onChange={(e) => setPrefs({ login_choice: e.target.value })}
            className="w-full rounded-lg border border-black/10 bg-white px-3 py-1.5 text-sm text-slate-900 outline-none focus:border-accent dark:border-white/10 dark:bg-surface-dark-alt dark:text-slate-100"
          >
            <option value="CPF">CPF</option>
            <option value="EMAIL">E-MAIL</option>
          </select>
        </div>
        <div>
          <label htmlFor="cadastro_fluxo" className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-200">
            Fluxo
          </label>
          <select
            id="cadastro_fluxo"
            aria-label="Escolher fluxo"
            title="Definir processo operacional (SELF ou FRONT)"
            value={prefs.fluxo}
            onChange={(e) => setPrefs({ fluxo: e.target.value })}
            className="w-full rounded-lg border border-black/10 bg-white px-3 py-1.5 text-sm text-slate-900 outline-none focus:border-accent dark:border-white/10 dark:bg-surface-dark-alt dark:text-slate-100"
          >
            <option value="SELF">SELF</option>
            <option value="FRONT">FRONT</option>
          </select>
        </div>
      </div>

      {cadastro.generating && (
        <div id="cadastro_progress" className="mt-4 h-2 overflow-hidden rounded-full bg-black/10 dark:bg-white/10">
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

      <button
        id="cadastro_btn"
        aria-label="Gerar cadastro"
        title="Processar a planilha e gerar arquivo tratado"
        disabled={cadastro.generating}
        onClick={handleSubmit}
        className="mt-4 rounded-lg bg-accent px-4 py-2 text-sm font-medium text-white hover:bg-accent-hover disabled:cursor-not-allowed disabled:opacity-50"
      >
        {cadastro.generating ? "Processando..." : "Gerar"}
      </button>

      <div id="cadastro_status" className="mt-3 text-sm" aria-live="polite">
        {cadastro.done && <span className="text-success">✔ Concluído</span>}
      </div>
      {cadastro.debugMsg && (
        <div id="cadastro_debug" className="mt-3 text-sm text-danger" aria-live="assertive">
          {cadastro.debugMsg}
        </div>
      )}
    </div>
  );
}
