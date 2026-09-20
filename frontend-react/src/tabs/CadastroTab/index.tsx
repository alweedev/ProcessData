import { useEffect, useRef, useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { cn } from "../../ui/cn";
import { FileList } from "../../ui/FileList";
import { IconAlert, IconCheck, IconUpload } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { ProcessingProgress } from "../../ui/ProcessingProgress";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { SegmentedControl } from "../../ui/SegmentedControl";
import { Stepper, type StepperItem } from "../../ui/Stepper";
import { TONE_OUTLINE } from "../../ui/tone";
import { ValidationReport } from "../../ui/ValidationReport";
import { MAX_FILES, OUTPUT_FILENAME, useCadastro } from "./useCadastro";
import { estadosDaLinhaDoTempo, etapaMaxima, podeIrPara, ULTIMA_ETAPA } from "./wizard";

const LOGIN_OPTIONS = [
  { value: "CPF", label: "CPF" },
  { value: "EMAIL", label: "E-mail" },
];

const FLUXO_OPTIONS = [
  { value: "SELF", label: "SELF" },
  { value: "FRONT", label: "FRONT" },
];

const ETAPAS = [
  { rotulo: "Fichas", titulo: "Enviar fichas", descricao: "Arraste ou selecione as planilhas de cadastro." },
  { rotulo: "Tipo de login", titulo: "Tipo de login", descricao: "Como os usuários vão acessar." },
  { rotulo: "Fluxo", titulo: "Fluxo", descricao: "Qual fluxo será usado neste cadastro." },
  { rotulo: "Gerar", titulo: "Validar e gerar", descricao: "Confira o resumo e gere o arquivo pronto para carga." },
];

interface Configuracao {
  login: string | null;
  fluxo: string | null;
}

/** Sem valor padrão e sem memória da última escolha (nem no localStorage): tipo de
 *  login e fluxo são escolhidos a cada cadastro, para não gerar com o da vez anterior. */
const SEM_CONFIGURACAO: Configuracao = { login: null, fluxo: null };

function plural(n: number, singular: string, pluralForma: string): string {
  return `${n} ${n === 1 ? singular : pluralForma}`;
}

function LinhaDoResumo({
  rotulo,
  valor,
  idAlterar,
  onAlterar,
  travado,
}: {
  rotulo: string;
  valor: string;
  idAlterar: string;
  onAlterar: () => void;
  travado: boolean;
}) {
  return (
    <li className="flex items-center gap-3 px-3 py-2.5 text-sm">
      <span className="w-28 shrink-0 text-text-muted">{rotulo}</span>
      <span className="min-w-0 flex-1 truncate font-medium text-text">{valor}</span>
      <button
        type="button"
        id={idAlterar}
        disabled={travado}
        onClick={onAlterar}
        className="rounded-control px-2 py-1 text-xs font-medium text-accent-text outline-none hover:underline focus-visible:ring-2 focus-visible:ring-accent/50 disabled:cursor-not-allowed disabled:opacity-50"
      >
        Alterar
      </button>
    </li>
  );
}

/**
 * Cadastro em massa como assistente de 4 etapas (fichas → tipo de login → fluxo → gerar),
 * guiado pela linha do tempo: uma etapa por vez, com o anel pulando de uma para a outra.
 */
export function CadastroTab() {
  const cadastro = useCadastro();
  const [config, setConfig] = useState<Configuracao>(SEM_CONFIGURACAO);
  const [etapa, setEtapa] = useState(0);
  const [sentido, setSentido] = useState<"fwd" | "back">("fwd");
  const [navegou, setNavegou] = useState(false);
  const [resetKey, setResetKey] = useState(0);
  const [helpOpen, setHelpOpen] = useState(false);
  const { login, fluxo } = config;

  const hasFiles = cadastro.files.length > 0;
  const finished = cadastro.done;
  const travado = cadastro.generating; // não se navega no meio do processamento
  const preenchido = [hasFiles, login !== null, fluxo !== null] as const;
  const estados = estadosDaLinhaDoTempo(etapa, preenchido, finished);
  const loginLabel = LOGIN_OPTIONS.find((o) => o.value === login)?.label ?? login;

  function irPara(destino: number) {
    if (destino === etapa) return;
    setSentido(destino > etapa ? "fwd" : "back");
    setNavegou(true);
    setEtapa(destino);
  }

  async function handleSubmit() {
    if (login === null || fluxo === null) return;
    await cadastro.submit(login, fluxo, () => {
      setResetKey((k) => k + 1);
      setConfig(SEM_CONFIGURACAO); // o próximo cadastro exige escolher de novo
    });
  }

  function handleValidate() {
    if (login === null || fluxo === null) return;
    void cadastro.validate(login, fluxo);
  }

  function limparFichas() {
    cadastro.clear();
    setResetKey((k) => k + 1);
  }

  function novoCadastro() {
    cadastro.clear();
    setConfig(SEM_CONFIGURACAO);
    setResetKey((k) => k + 1);
    irPara(0);
  }

  // Ao trocar de etapa, leva o foco ao título: quem usa teclado ou leitor de tela
  // começa a nova etapa do começo, em vez de ficar num botão que sumiu.
  const tituloRef = useRef<HTMLHeadingElement>(null);
  const primeiraRender = useRef(true);
  useEffect(() => {
    if (primeiraRender.current) {
      primeiraRender.current = false;
      return;
    }
    tituloRef.current?.focus({ preventScroll: true });
  }, [etapa]);

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

  const passos: StepperItem[] = [
    {
      label: ETAPAS[0].rotulo,
      detail: finished ? "Enviadas" : hasFiles ? plural(cadastro.files.length, "arquivo", "arquivos") : "Planilhas .xlsx ou .xls",
    },
    { label: ETAPAS[1].rotulo, detail: finished ? "Definido" : (loginLabel ?? "Escolha uma opção") },
    { label: ETAPAS[2].rotulo, detail: finished ? "Definido" : (fluxo ?? "Escolha uma opção") },
    { label: ETAPAS[3].rotulo, detail: finished ? "Arquivo gerado" : "Validar e baixar" },
  ].map((passo, i) => ({
    ...passo,
    state: estados[i],
    selectable: !travado && podeIrPara(i, preenchido, finished),
  }));

  // Depois de gerar, o cabeçalho deixa de pedir para "conferir o resumo".
  const cabecalho = finished
    ? { titulo: "Tudo pronto", descricao: "Comece outro cadastro quando quiser." }
    : ETAPAS[etapa];

  const nomeDasFichas =
    cadastro.files.length === 1 ? cadastro.files[0].name : plural(cadastro.files.length, "arquivo", "arquivos");

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
            <strong>1º Passo:</strong> carregue a planilha Excel (<strong>.xlsx/.xls</strong>) já preenchida e clique em{" "}
            <em>Continuar</em>.
          </li>
          <li>
            <strong>2º Passo:</strong> escolha o <strong>tipo de login</strong> (CPF ou e-mail). Ao escolher, o
            assistente avança sozinho.
          </li>
          <li>
            <strong>3º Passo:</strong> escolha o <strong>fluxo</strong> (SELF ou FRONT). Essas escolhas não ficam
            salvas: é preciso informá-las a cada cadastro.
          </li>
          <li>
            <strong>4º Passo:</strong> confira o resumo, valide se quiser e clique em <em>Gerar cadastro</em>.
            Recupere execuções no <strong>Histórico</strong> quando precisar.
          </li>
        </ol>
        <p className="mt-2 text-text-subtle">
          Dica: valide a planilha antes de gerar para evitar retrabalho. Os pontos já concluídos da linha do tempo
          levam de volta à etapa.
        </p>
      </Modal>

      <Stepper label="Etapas do cadastro" steps={passos} onSelect={irPara} />

      <div className="grid items-start gap-6 xl:grid-cols-[minmax(0,1fr)_22rem]">
        <Card padding="lg">
          <div
            key={etapa}
            className={cn(
              navegou && (sentido === "fwd" ? "animate-[pd-step-in-fwd_280ms_ease-out]" : "animate-[pd-step-in-back_280ms_ease-out]"),
            )}
          >
            <header className="mb-4">
              <p className="text-xs font-medium text-text-muted">
                Etapa {etapa + 1} de {ULTIMA_ETAPA + 1}
              </p>
              <h3
                ref={tituloRef}
                id="cadastro_step_title"
                tabIndex={-1}
                className="text-lg font-semibold text-text outline-none"
              >
                {cabecalho.titulo}
              </h3>
              <p className="mt-0.5 text-sm text-text-muted">{cabecalho.descricao}</p>
            </header>

            {etapa === 0 && (
              <>
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
                  description={
                    hasFiles
                      ? "Arraste outras planilhas aqui para trocar a seleção"
                      : "Arraste as planilhas aqui ou clique para selecionar"
                  }
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
                    onClearAll={limparFichas}
                    clearAllId="cadastro_clear_btn"
                  />
                </div>
              </>
            )}

            {etapa === 1 && (
              <SegmentedControl
                id="cadastro_login_choice"
                label="Tipo de login"
                size="lg"
                value={login}
                options={LOGIN_OPTIONS}
                onChange={(value) => setConfig((c) => ({ ...c, login: value }))}
                onCommit={() => irPara(etapaMaxima([hasFiles, true, fluxo !== null]))}
                required
              />
            )}

            {etapa === 2 && (
              <SegmentedControl
                id="cadastro_fluxo"
                label="Fluxo"
                size="lg"
                value={fluxo}
                options={FLUXO_OPTIONS}
                onChange={(value) => setConfig((c) => ({ ...c, fluxo: value }))}
                onCommit={() => irPara(etapaMaxima([hasFiles, login !== null, true]))}
                required
              />
            )}

            {etapa === ULTIMA_ETAPA && (
              <>
                {!finished && (
                  <>
                    <ul className="mb-4 divide-y divide-border rounded-control border border-border">
                      <LinhaDoResumo
                        rotulo="Fichas"
                        valor={nomeDasFichas}
                        idAlterar="cadastro_edit_fichas"
                        onAlterar={() => irPara(0)}
                        travado={travado}
                      />
                      <LinhaDoResumo
                        rotulo="Tipo de login"
                        valor={loginLabel ?? ""}
                        idAlterar="cadastro_edit_login"
                        onAlterar={() => irPara(1)}
                        travado={travado}
                      />
                      <LinhaDoResumo
                        rotulo="Fluxo"
                        valor={fluxo ?? ""}
                        idAlterar="cadastro_edit_fluxo"
                        onAlterar={() => irPara(2)}
                        travado={travado}
                      />
                    </ul>

                    <div className="flex flex-wrap items-center gap-3">
                      <Button
                        id="cadastro_validate_btn"
                        variant="secondary"
                        disabled={travado}
                        loading={cadastro.validation.status === "loading"}
                        onClick={handleValidate}
                      >
                        {cadastro.validation.status === "loading" ? "Validando..." : "Validar planilha"}
                      </Button>
                      <Button
                        id="cadastro_btn"
                        aria-label="Gerar cadastro"
                        title="Processar a planilha e gerar arquivo tratado"
                        loading={cadastro.generating}
                        onClick={handleSubmit}
                      >
                        {cadastro.generating ? "Processando..." : "Gerar cadastro"}
                      </Button>
                    </div>
                  </>
                )}

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
                    {finished && (
                      <div className="flex items-start gap-3 rounded-control border border-success/40 bg-success-soft px-4 py-3">
                        <span className="mt-0.5 text-success">
                          <IconCheck className="h-5 w-5" />
                        </span>
                        <div className="text-sm">
                          <p className="font-medium text-success">Cadastro concluído</p>
                          <p className="mt-0.5 text-text-muted">
                            O arquivo <strong className="font-medium text-text">{OUTPUT_FILENAME}</strong> foi gerado.
                            Se o navegador não o salvou, baixe de novo em “Nesta sessão”.
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
              </>
            )}
          </div>

          <div className="mt-6 flex items-center justify-between gap-3 border-t border-border pt-4">
            {etapa > 0 && !finished ? (
              <Button id="cadastro_back_btn" variant="ghost" disabled={travado} onClick={() => irPara(etapa - 1)}>
                Voltar
              </Button>
            ) : (
              <span />
            )}
            {finished ? (
              <Button id="cadastro_new_btn" onClick={novoCadastro}>
                Novo cadastro
              </Button>
            ) : (
              etapa < ULTIMA_ETAPA && (
                <Button id="cadastro_next_btn" disabled={!preenchido[etapa]} onClick={() => irPara(etapa + 1)}>
                  Continuar
                </Button>
              )
            )}
          </div>
        </Card>

        <div className="xl:sticky xl:top-20">
          <RunHistoryPanel operation="cadastro" className="mt-0" />
        </div>
      </div>
    </div>
  );
}
