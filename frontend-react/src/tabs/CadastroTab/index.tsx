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
    <div className="tw:rounded-xl tw:border tw:border-black/10 tw:bg-surface tw:p-4 tw:shadow-sm tw:dark:border-white/10 tw:dark:bg-surface-dark tw:sm:p-6">
      <div className="tw:mb-3 tw:flex tw:flex-wrap tw:items-center tw:justify-between tw:gap-2">
        <h2 className="tw:text-lg tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">Cadastro Carga</h2>
        <button
          id="cadastro_help_btn"
          type="button"
          aria-label="Como usar o cadastro em lote"
          title="Guia do cadastro em lote"
          onClick={() => setHelpOpen(true)}
          className="tw:rounded-lg tw:border tw:border-info tw:px-3 tw:py-1.5 tw:text-sm tw:font-medium tw:text-info tw:hover:bg-info/10"
        >
          Como usar
        </button>
      </div>

      <Modal open={helpOpen} onClose={() => setHelpOpen(false)} title="Como usar carga cadastro">
        <ol className="tw:list-decimal tw:space-y-1 tw:pl-5">
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
        <p className="tw:mt-2 tw:text-slate-500 tw:dark:text-slate-400">
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
        <div id="cadastro_uploadFeedback" aria-live="polite" className="tw:mt-3">
          {cadastro.files.length > 0 && (
            <span className="tw:inline-flex tw:items-center tw:gap-2 tw:rounded-full tw:bg-success/10 tw:px-3 tw:py-1 tw:text-sm tw:text-success">
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
                className="tw:text-success tw:hover:text-danger"
              >
                ✕
              </button>
            </span>
          )}
        </div>
      </FileDropzone>

      <div className="tw:mt-4 tw:grid tw:gap-3 tw:sm:grid-cols-2">
        <div>
          <label htmlFor="cadastro_login_choice" className="tw:mb-1 tw:block tw:text-sm tw:font-medium tw:text-slate-700 tw:dark:text-slate-200">
            Tipo de login
          </label>
          <select
            id="cadastro_login_choice"
            aria-label="Escolher tipo de login"
            title="Selecionar formato principal de acesso"
            value={prefs.login_choice}
            onChange={(e) => setPrefs({ login_choice: e.target.value })}
            className="tw:w-full tw:rounded-lg tw:border tw:border-black/10 tw:bg-white tw:px-3 tw:py-1.5 tw:text-sm tw:text-slate-900 tw:outline-none tw:focus:border-accent tw:dark:border-white/10 tw:dark:bg-surface-dark-alt tw:dark:text-slate-100"
          >
            <option value="CPF">CPF</option>
            <option value="EMAIL">E-MAIL</option>
          </select>
        </div>
        <div>
          <label htmlFor="cadastro_fluxo" className="tw:mb-1 tw:block tw:text-sm tw:font-medium tw:text-slate-700 tw:dark:text-slate-200">
            Fluxo
          </label>
          <select
            id="cadastro_fluxo"
            aria-label="Escolher fluxo"
            title="Definir processo operacional (SELF ou FRONT)"
            value={prefs.fluxo}
            onChange={(e) => setPrefs({ fluxo: e.target.value })}
            className="tw:w-full tw:rounded-lg tw:border tw:border-black/10 tw:bg-white tw:px-3 tw:py-1.5 tw:text-sm tw:text-slate-900 tw:outline-none tw:focus:border-accent tw:dark:border-white/10 tw:dark:bg-surface-dark-alt tw:dark:text-slate-100"
          >
            <option value="SELF">SELF</option>
            <option value="FRONT">FRONT</option>
          </select>
        </div>
      </div>

      {cadastro.generating && (
        <div id="cadastro_progress" className="tw:mt-4 tw:h-2 tw:overflow-hidden tw:rounded-full tw:bg-black/10 tw:dark:bg-white/10">
          <div
            id="cadastro_progressBar"
            role="progressbar"
            aria-valuenow={Math.round(cadastro.progress)}
            aria-valuemin={0}
            aria-valuemax={100}
            style={{ width: `${cadastro.progress}%` }}
            className="tw:h-full tw:bg-accent tw:transition-[width]"
          />
        </div>
      )}

      <button
        id="cadastro_btn"
        aria-label="Gerar cadastro"
        title="Processar a planilha e gerar arquivo tratado"
        disabled={cadastro.generating}
        onClick={handleSubmit}
        className="tw:mt-4 tw:rounded-lg tw:bg-accent tw:px-4 tw:py-2 tw:text-sm tw:font-medium tw:text-white tw:hover:bg-accent-hover tw:disabled:cursor-not-allowed tw:disabled:opacity-50"
      >
        {cadastro.generating ? "Processando..." : "Gerar"}
      </button>

      <div id="cadastro_status" className="tw:mt-3 tw:text-sm" aria-live="polite">
        {cadastro.done && <span className="tw:text-success">✔ Concluído</span>}
      </div>
      {cadastro.debugMsg && (
        <div id="cadastro_debug" className="tw:mt-3 tw:text-sm tw:text-danger" aria-live="assertive">
          {cadastro.debugMsg}
        </div>
      )}
    </div>
  );
}
