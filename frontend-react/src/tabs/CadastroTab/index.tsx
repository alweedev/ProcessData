import { useEffect, useRef, useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { cn } from "../../ui/cn";
import { FileList } from "../../ui/FileList";
import { IconCheck, IconUpload } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { ProcessingProgress } from "../../ui/ProcessingProgress";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { SegmentedControl } from "../../ui/SegmentedControl";
import type { NameReviewItem } from "../../lib/api";
import { acceptAll, buildOverrides, pendingItems, type Decisions, type NameDecision } from "../../lib/nameReview";
import { explainFailure } from "../../lib/failure";
import { pendenciasDoRelatorio } from "../../lib/qualityReport";
import { GenerationError } from "../../ui/GenerationError";
import { NamesReview } from "../../ui/NamesReview";
import { Stepper, type StepperItem } from "../../ui/Stepper";
import { ValidationReport } from "../../ui/ValidationReport";
import { ValidationSummary } from "../../ui/ValidationSummary";
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
  {
    rotulo: "Fichas",
    titulo: "Enviar fichas",
    descricao: "Arraste ou selecione as planilhas de cadastro.",
  },
  {
    rotulo: "Tipo de login",
    titulo: "Tipo de login",
    descricao: "Como os usuários vão acessar.",
  },
  {
    rotulo: "Fluxo",
    titulo: "Fluxo",
    descricao: "Qual fluxo será usado neste cadastro.",
  },
  {
    rotulo: "Gerar",
    titulo: "Validar e gerar",
    descricao: "Confira o resumo e gere o arquivo pronto para carga.",
  },
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
  acao = "Alterar",
  onAlterar,
  travado,
}: {
  rotulo: string;
  valor: string;
  idAlterar: string;
  /** o que o botão diz: para as fichas, "Trocar arquivos" (abre a escolha ali mesmo) */
  acao?: string;
  onAlterar: () => void;
  travado: boolean;
}) {
  return (
    <li className="flex items-center gap-3 px-3 py-2.5 text-sm">
      <span className="w-28 shrink-0 text-text-muted">{rotulo}</span>
      <span className="min-w-0 flex-1 truncate font-medium text-text">{valor}</span>
      <Button id={idAlterar} variant="outline" size="sm" disabled={travado} onClick={onAlterar}>
        {acao}
      </Button>
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

  const [confirmando, setConfirmando] = useState(false);
  const validando = cadastro.validation.status === "loading";
  // O servidor recusou a ficha (ilegível, sem colunas conhecidas, vazia): gerar falharia pelo mesmo motivo.
  const planilhaRecusada = Object.keys(cadastro.validation.failure?.fileErrors ?? {}).length > 0;
  const relatorio = cadastro.validation.status === "done" ? cadastro.validation.report : null;
  const { graves, avisos } = relatorio ? pendenciasDoRelatorio(relatorio) : { graves: [], avisos: [] };

  // Conferência de nomes: o que o usuário decidiu para cada nome duvidoso (ou acima de 20 caracteres). As decisões valem
  // para UMA validação — a lista muda de identidade quando a validação recomeça, e aí elas são descartadas.
  const nameReview = cadastro.validation.nameReview;
  const [decisionState, setDecisionState] = useState<{
    review: NameReviewItem[];
    decisions: Decisions;
  }>({
    review: nameReview,
    decisions: {},
  });
  const decisoes = decisionState.review === nameReview ? decisionState.decisions : {};
  const nomesPendentes = pendingItems(nameReview, decisoes).length;

  function decidirNome(key: string, decision: NameDecision) {
    setDecisionState({
      review: nameReview,
      decisions: { ...decisoes, [key]: decision },
    });
  }

  function aceitarTodasAsSugestoes() {
    setDecisionState({
      review: nameReview,
      decisions: acceptAll(nameReview, decisoes),
    });
  }

  async function gerar() {
    setConfirmando(false);
    if (login === null || fluxo === null) return;
    await cadastro.submit(
      login,
      fluxo,
      () => {
        setResetKey((k) => k + 1);
        setConfig(SEM_CONFIGURACAO); // o próximo cadastro exige escolher de novo
      },
      buildOverrides(nameReview, decisoes),
    );
  }

  /** Com pendências graves apontadas pela validação, gerar exige confirmação; sem elas (ou só com avisos), gera direto. */
  function handleSubmit() {
    if (graves.length > 0) setConfirmando(true);
    else void gerar();
  }

  // Ao chegar em "Gerar" a planilha é validada sozinha: o veredito já está à vista, sem depender de
  // o usuário lembrar do botão. Mudou algo (fichas, login, fluxo)? A validação foi descartada e roda de novo.
  const validarSozinha =
    etapa === ULTIMA_ETAPA &&
    !finished &&
    !travado &&
    hasFiles &&
    login !== null &&
    fluxo !== null &&
    cadastro.validation.status === "idle";
  const { validate } = cadastro;
  useEffect(() => {
    if (validarSozinha && login !== null && fluxo !== null) void validate(login, fluxo);
  }, [validarSozinha, login, fluxo, validate]);

  /** Sem fichas o que veio depois perde o sentido: recomeça a escolha de login e fluxo, para a
   *  linha do tempo e o resumo não seguirem mostrando como certo algo sem base. */
  function limparFichas() {
    cadastro.clear();
    setConfig(SEM_CONFIGURACAO);
    setResetKey((k) => k + 1);
  }

  /** Trocar as fichas sem sair da última etapa: escolhe os arquivos ali mesmo; login e fluxo ficam como estão e a
   *  validação roda de novo sozinha. */
  const trocarFichasInput = useRef<HTMLInputElement>(null);
  function trocarFichas(list: FileList | null) {
    if (list && list.length > 0) cadastro.pickFiles(list);
    if (trocarFichasInput.current) trocarFichasInput.current.value = ""; // permite escolher o mesmo arquivo de novo
  }

  function removerFicha(index: number) {
    cadastro.removeFile(index);
    if (cadastro.files.length === 1) setConfig(SEM_CONFIGURACAO); // era a última
  }

  /** O relatório de validação vale só para o login e o fluxo com que foi rodado. */
  function escolher(campo: keyof Configuracao, valor: string) {
    if (config[campo] === valor) return;
    setConfig((c) => ({ ...c, [campo]: valor }));
    cadastro.invalidateValidation();
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

  // O sucesso/erro da geração nasce abaixo dos botões, muitas vezes fora da tela: leva à vista quando ele muda. (O
  // veredito da validação não precisa: roda sozinha ao chegar em "Gerar" e fica acima dos botões.)
  const resultRef = useRef<HTMLDivElement>(null);
  const resultKey = cadastro.done ? "sucesso" : cadastro.failure ? "erro" : null;
  useEffect(() => {
    if (!resultKey) return;
    const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
    resultRef.current?.scrollIntoView({
      block: "nearest",
      behavior: reduced ? "auto" : "smooth",
    });
  }, [resultKey]);

  const passos: StepperItem[] = [
    {
      label: ETAPAS[0].rotulo,
      detail: finished
        ? "Enviadas"
        : hasFiles
          ? plural(cadastro.files.length, "arquivo", "arquivos")
          : "Planilhas .xlsx ou .xls",
    },
    {
      label: ETAPAS[1].rotulo,
      detail: finished ? "Definido" : (loginLabel ?? "Escolha uma opção"),
    },
    {
      label: ETAPAS[2].rotulo,
      detail: finished ? "Definido" : (fluxo ?? "Escolha uma opção"),
    },
    {
      label: ETAPAS[3].rotulo,
      detail: finished ? "Arquivo gerado" : "Validar e baixar",
    },
  ].map((passo, i) => ({
    ...passo,
    state: estados[i],
    selectable: !travado && podeIrPara(i, preenchido, finished),
  }));

  // Depois de gerar, o cabeçalho deixa de pedir para "conferir o resumo".
  const cabecalho = finished
    ? {
        titulo: "Tudo pronto",
        descricao: "Comece outro cadastro quando quiser.",
      }
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
            <strong>4º Passo:</strong> confira o resumo, valide se quiser e clique em <em>Gerar cadastro</em>. Recupere
            execuções no <strong>Histórico</strong> quando precisar.
          </li>
        </ol>
        <p className="mt-2 text-text-subtle">
          Dica: valide a planilha antes de gerar para evitar retrabalho. Os pontos já concluídos da linha do tempo levam
          de volta à etapa.
        </p>
      </Modal>

      <Modal
        open={confirmando}
        onClose={() => setConfirmando(false)}
        title="Gerar mesmo com pendências?"
        footer={
          <>
            <Button id="cadastro_confirm_cancel_btn" variant="ghost" onClick={() => setConfirmando(false)}>
              Revisar antes
            </Button>
            <Button id="cadastro_confirm_generate_btn" onClick={() => void gerar()}>
              Gerar mesmo assim
            </Button>
          </>
        }
      >
        <p>A validação encontrou pendências nesta planilha:</p>
        <ul className="my-2 list-disc space-y-0.5 pl-5 text-text">
          {[...graves, ...avisos].map((pendencia) => (
            <li key={pendencia}>{pendencia}</li>
          ))}
        </ul>
        <p>Confira o relatório antes de gerar, ou gere assim mesmo se estiver ciente.</p>
      </Modal>

      <Stepper label="Etapas do cadastro" steps={passos} onSelect={irPara} />

      <div className="grid items-start gap-6 xl:grid-cols-[minmax(0,1fr)_22rem]">
        <Card padding="lg">
          <div
            key={etapa}
            className={cn(
              navegou &&
                (sentido === "fwd"
                  ? "animate-[pd-step-in-fwd_280ms_ease-out]"
                  : "animate-[pd-step-in-back_280ms_ease-out]"),
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
                    // Recusada: o que já estava escolhido fica; só o <input> nativo, com a lista recusada, é refeito.
                    if (!accepted) setResetKey((k) => k + 1);
                  }}
                />
                <div id="cadastro_uploadFeedback" aria-live="polite">
                  <FileList
                    files={cadastro.files}
                    maxFiles={MAX_FILES}
                    onRemove={removerFicha}
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
                onChange={(value) => escolher("login", value)}
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
                onChange={(value) => escolher("fluxo", value)}
                onCommit={() => irPara(etapaMaxima([hasFiles, login !== null, true]))}
                required
              />
            )}

            {etapa === ULTIMA_ETAPA && (
              <>
                {!finished && (
                  <>
                    <input
                      ref={trocarFichasInput}
                      id="cadastro_swap_files"
                      type="file"
                      accept=".xlsx,.xls"
                      multiple
                      hidden
                      aria-label="Escolher outras fichas"
                      onChange={(e) => trocarFichas(e.target.files)}
                    />
                    <ul className="mb-4 divide-y divide-border rounded-control border border-border">
                      <LinhaDoResumo
                        rotulo="Fichas"
                        valor={nomeDasFichas}
                        idAlterar="cadastro_edit_fichas"
                        acao="Trocar arquivos"
                        onAlterar={() => trocarFichasInput.current?.click()}
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

                    {/* O resultado vem ANTES dos botões: o veredito com o que corrigir num cartão só, e os nomes a conferir. */}
                    <div>
                      <ValidationSummary status={cadastro.validation.status} report={relatorio}>
                        <ValidationReport report={relatorio} />
                      </ValidationSummary>

                      {nameReview.length > 0 && (
                        <NamesReview
                          items={nameReview}
                          decisions={decisoes}
                          onDecision={decidirNome}
                          onAcceptAll={aceitarTodasAsSugestoes}
                        />
                      )}

                      {cadastro.validation.failure ? (
                        <GenerationError
                          id="cadastro_validation_error"
                          className="mb-4"
                          view={explainFailure(cadastro.validation.failure, "Não foi possível validar a planilha")}
                        />
                      ) : (
                        cadastro.validation.status === "error" && (
                          <ValidationReport report={null} error={cadastro.validation.error} />
                        )
                      )}
                    </div>

                    <div className="flex flex-wrap items-center gap-3">
                      <Button
                        id="cadastro_btn"
                        // Com pendência, gerar segue à vista, em âmbar (pede confirmação): destaque sem fingir que está tudo certo.
                        variant={graves.length > 0 ? "warning" : "primary"}
                        aria-label="Gerar cadastro"
                        title={
                          validando
                            ? "Aguarde a validação terminar"
                            : planilhaRecusada
                              ? "A planilha foi recusada: troque a ficha em “Alterar”"
                              : nomesPendentes > 0
                                ? "Confira os nomes destacados antes de gerar"
                                : "Processar a planilha e gerar arquivo tratado"
                        }
                        disabled={validando || planilhaRecusada || nomesPendentes > 0}
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

                  {cadastro.failure && <GenerationError className="mt-4" view={explainFailure(cadastro.failure)} />}
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
