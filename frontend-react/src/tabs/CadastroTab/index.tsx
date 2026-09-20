import { useEffect, useRef, useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { usePersistedState } from "../../hooks/usePersistedState";
import { Badge } from "../../ui/Badge";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { FileList } from "../../ui/FileList";
import { IconAlert, IconCheck, IconUpload } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { ProcessingProgress } from "../../ui/ProcessingProgress";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { SegmentedControl } from "../../ui/SegmentedControl";
import { StepSection } from "../../ui/StepSection";
import { Stepper, type StepperItem } from "../../ui/Stepper";
import { TONE_OUTLINE } from "../../ui/tone";
import { ValidationReport } from "../../ui/ValidationReport";
import { MAX_FILES, useCadastro } from "./useCadastro";

const LOGIN_OPTIONS = [
  { value: "CPF", label: "CPF" },
  { value: "EMAIL", label: "E-mail" },
];

const FLUXO_OPTIONS = [
  { value: "SELF", label: "SELF" },
  { value: "FRONT", label: "FRONT" },
];

export function CadastroTab() {
  const cadastro = useCadastro();
  const [prefs, setPrefs] = usePersistedState("cadastro_prefs", { login_choice: "CPF", fluxo: "SELF" });
  const [resetKey, setResetKey] = useState(0);
  const [helpOpen, setHelpOpen] = useState(false);

  async function handleSubmit() {
    await cadastro.submit(prefs.login_choice, prefs.fluxo, () => setResetKey((k) => k + 1));
  }

  function clear() {
    cadastro.clear();
    setResetKey((k) => k + 1);
  }

  const hasFiles = cadastro.files.length > 0;

  // O relatório/sucesso/erro nascem abaixo dos botões, muitas vezes fora da
  // tela; leva o resultado à vista quando ele muda.
  const resultRef = useRef<HTMLDivElement>(null);
  const validationStatus = cadastro.validation.status;
  const resultKey =
    validationStatus === "done" || validationStatus === "error"
      ? "validacao"
      : cadastro.done
        ? "sucesso"
        : cadastro.debugMsg
          ? "erro"
          : null;
  useEffect(() => {
    if (!resultKey) return;
    const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
    resultRef.current?.scrollIntoView({ block: "nearest", behavior: reduced ? "auto" : "smooth" });
  }, [resultKey]);
  const loginLabel = LOGIN_OPTIONS.find((o) => o.value === prefs.login_choice)?.label ?? prefs.login_choice;

  const steps: StepperItem[] = [
    {
      label: "Enviar fichas",
      detail: hasFiles ? `${cadastro.files.length} ${cadastro.files.length === 1 ? "arquivo" : "arquivos"}` : "Planilhas .xlsx ou .xls",
      state: hasFiles || cadastro.done ? "done" : "current",
    },
    {
      label: "Configurar",
      detail: `${loginLabel} · ${prefs.fluxo}`,
      state: hasFiles || cadastro.done ? "done" : "todo",
    },
    {
      label: "Validar e gerar",
      detail: cadastro.done ? "Arquivo gerado" : "Arquivo pronto para carga",
      state: cadastro.done ? "done" : hasFiles ? "current" : "todo",
    },
  ];

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

      <Stepper label="Etapas do cadastro" steps={steps} />

      <div className="grid items-start gap-6 xl:grid-cols-[minmax(0,1fr)_22rem]">
        <Card padding="lg">
          <StepSection number={1} title="Enviar fichas" description="arraste ou selecione as planilhas">
            <FileDropzone
              key={resetKey}
              variant="rich"
              compact={hasFiles}
              buttonLabel={hasFiles ? "Trocar arquivos" : "Selecionar"}
              id="cadastro_files"
              containerId="cadastro_uploadArea"
              accept=".xlsx,.xls"
              multiple
              ariaLabel="Upload da base de cadastro. Pressione para selecionar arquivo"
              description={hasFiles ? "Arraste outras planilhas aqui para trocar a seleção" : "Arraste as planilhas aqui ou clique para selecionar"}
              hint={`Excel .xlsx ou .xls · até ${MAX_FILES} arquivos · 10 MB cada`}
              syncFiles={cadastro.files}
              onFiles={(list) => {
                const accepted = cadastro.pickFiles(list);
                if (!accepted) setResetKey((k) => k + 1);
              }}
            />
            <div id="cadastro_uploadFeedback" aria-live="polite">
              <FileList
                files={cadastro.files}
                maxFiles={MAX_FILES}
                onRemove={cadastro.removeFile}
                onClearAll={clear}
                clearAllId="cadastro_clear_btn"
              />
            </div>
          </StepSection>

          <StepSection number={2} title="Configurar" description="como o arquivo de carga será montado">
            <div className="grid gap-4 sm:grid-cols-2">
              <SegmentedControl
                id="cadastro_login_choice"
                label="Tipo de login"
                value={prefs.login_choice}
                options={LOGIN_OPTIONS}
                onChange={(value) => setPrefs({ login_choice: value })}
              />
              <SegmentedControl
                id="cadastro_fluxo"
                label="Fluxo"
                value={prefs.fluxo}
                options={FLUXO_OPTIONS}
                onChange={(value) => setPrefs({ fluxo: value })}
              />
            </div>
          </StepSection>

          <StepSection number={3} title="Validar e gerar" description="confira antes de gerar o arquivo">
            <div className="flex flex-wrap items-center gap-3">
              <Button
                id="cadastro_validate_btn"
                variant="secondary"
                disabled={!hasFiles || cadastro.generating}
                aria-describedby={hasFiles ? undefined : "cadastro_hint"}
                loading={cadastro.validation.status === "loading"}
                onClick={() => cadastro.validate(prefs.login_choice, prefs.fluxo)}
              >
                {cadastro.validation.status === "loading" ? "Validando..." : "Validar planilha"}
              </Button>
              <Button
                id="cadastro_btn"
                aria-label="Gerar cadastro"
                title="Processar a planilha e gerar arquivo tratado"
                disabled={!hasFiles}
                aria-describedby={hasFiles ? undefined : "cadastro_hint"}
                loading={cadastro.generating}
                onClick={handleSubmit}
              >
                {cadastro.generating ? "Processando..." : "Gerar cadastro"}
              </Button>
              {hasFiles ? (
                <span className="flex flex-wrap items-center gap-1.5 text-xs text-text-muted">
                  Pronto para gerar:
                  <Badge tone="neutral">
                    {cadastro.files.length} {cadastro.files.length === 1 ? "arquivo" : "arquivos"}
                  </Badge>
                  <Badge tone="neutral">Login {loginLabel}</Badge>
                  <Badge tone="neutral">Fluxo {prefs.fluxo}</Badge>
                </span>
              ) : (
                <span id="cadastro_hint" className="text-xs text-text-muted">
                  Envie ao menos uma ficha para validar e gerar.
                </span>
              )}
            </div>

            {cadastro.generating && (
              <div className="mt-4">
                <ProcessingProgress id="cadastro_progress" progress={cadastro.progress} />
              </div>
            )}

            <div ref={resultRef} className="scroll-mt-16">
              {cadastro.validation.status !== "idle" && (
                <div className="mt-4">
                  <ValidationReport
                    report={cadastro.validation.report}
                    loading={cadastro.validation.status === "loading"}
                    error={cadastro.validation.status === "error" ? cadastro.validation.error : null}
                  />
                </div>
              )}

              <div id="cadastro_status" aria-live="polite">
                {cadastro.done && (
                  <div className="mt-4 flex items-start gap-3 rounded-control border border-success/40 bg-success-soft px-4 py-3">
                    <span className="mt-0.5 text-success">
                      <IconCheck className="h-5 w-5" />
                    </span>
                    <div className="text-sm">
                      <p className="font-medium text-success">Cadastro gerado com sucesso</p>
                      <p className="mt-0.5 text-text-muted">
                        O download de <strong className="font-medium text-text">saida_cadastro.xlsx</strong> começou. O
                        arquivo também fica em “Nesta sessão”, para baixar de novo.
                      </p>
                    </div>
                  </div>
                )}
              </div>

              {cadastro.debugMsg && (
                <div
                  id="cadastro_debug"
                  aria-live="assertive"
                  className={`mt-4 flex items-start gap-3 rounded-control border px-4 py-3 text-sm ${TONE_OUTLINE.danger}`}
                >
                  <IconAlert className="mt-0.5 h-5 w-5 shrink-0" />
                  <span>{cadastro.debugMsg}</span>
                </div>
              )}
            </div>
          </StepSection>
        </Card>

        <div className="xl:sticky xl:top-20">
          <RunHistoryPanel operation="cadastro" className="mt-0" />
        </div>
      </div>
    </div>
  );
}
