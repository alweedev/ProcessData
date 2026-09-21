import { useEffect, useRef, useState } from "react";
import { FileDropzone } from "../../components/FileDropzone";
import { Modal } from "../../components/Modal";
import { explainFailure } from "../../lib/failure";
import { Button } from "../../ui/Button";
import { Card } from "../../ui/Card";
import { Field } from "../../ui/Field";
import { GenerationError } from "../../ui/GenerationError";
import { IconCheck, IconUserMinus } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { ProcessingProgress } from "../../ui/ProcessingProgress";
import { RunHistoryPanel } from "../../ui/RunHistoryPanel";
import { Stepper, type StepperItem } from "../../ui/Stepper";
import { Textarea } from "../../ui/Textarea";
import { ImpactoCard } from "./ImpactoCard";
import { ARQUIVO_ZIP, useInativacao } from "./useInativacao";
import { podeContinuar, podeExecutar, type Confirmacao } from "./wizard";

const ETAPAS = [
  {
    rotulo: "Bases e lista",
    titulo: "Enviar bases e lista",
    descricao: "Envie a base de cadastro e a de estruturas de aprovação e informe quem será inativado.",
  },
  {
    rotulo: "Impacto",
    titulo: "Conferir o impacto",
    descricao: "Veja o que muda antes de aplicar. Nada é alterado até você confirmar.",
  },
  {
    rotulo: "Confirmar",
    titulo: "Confirmar e executar",
    descricao: "Confirme para inativar os usuários e gerar os arquivos para carga.",
  },
];

const SEM_CONFIRMACAO = { digital: "", impacto: false, orfas: false };

function ArquivoEscolhido({
  file,
  id,
  disabled,
  onClear,
}: {
  file: File | null;
  id: string;
  disabled: boolean;
  onClear: () => void;
}) {
  return (
    <div id={`${id}_feedback`} aria-live="polite" className="mt-3">
      {file && (
        <span className="inline-flex items-center gap-2 rounded-pill border border-border bg-surface-2 py-1 pl-3 pr-1.5 text-sm text-text">
          {file.name}
          <button
            id={`${id}_clear_btn`}
            type="button"
            title="Remover arquivo"
            aria-label={`Remover ${file.name}`}
            disabled={disabled}
            onClick={(e) => {
              e.stopPropagation();
              onClear();
            }}
            className="flex h-5 w-5 items-center justify-center rounded-pill text-text-subtle transition-colors hover:bg-danger/10 hover:text-danger disabled:cursor-not-allowed disabled:opacity-50 disabled:hover:bg-transparent disabled:hover:text-text-subtle"
          >
            ✕
          </button>
        </span>
      )}
    </div>
  );
}

function plural(n: number, singular: string, pluralForma: string): string {
  return `${n} ${n === 1 ? singular : pluralForma}`;
}

/**
 * Inativação em cascata como assistente de 3 etapas: bases e lista → impacto → confirmar.
 * A análise não altera nada; só a confirmação na última etapa dispara a execução.
 */
export function InativacaoTab() {
  const inativacao = useInativacao();
  const { analise, concluido, executando, analisando } = inativacao;
  const [etapa, setEtapa] = useState(0);
  const [helpOpen, setHelpOpen] = useState(false);
  const [confirmacao, setConfirmacao] = useState(SEM_CONFIRMACAO);

  const digital = analise?.impressaoDigital ?? "";
  // A confirmação vale só para a análise em que foi dada: outra análise, outra confirmação.
  const confirmada: Confirmacao = confirmacao.digital === digital ? confirmacao : SEM_CONFIRMACAO;
  // Concluída, só o cartão de sucesso importa (a última etapa), aconteça o que acontecer com as entradas.
  // Sem análise válida (nada analisado ou ela foi descartada) só a primeira etapa faz sentido.
  const etapaAtual = concluido ? ETAPAS.length - 1 : analise ? etapa : 0;
  const resumo = analise?.resumo;
  const orfas = resumo?.estruturasOrfas ?? 0;
  const pendentes = analise ? analise.usuarios.filter((u) => u.situacao === "PENDENTE_SELECAO").length : 0;
  // Analisando ou executando: nada que mude as entradas ou troque de etapa fica ao alcance.
  const travado = analisando || executando;

  function marcar(campo: keyof Confirmacao, valor: boolean) {
    setConfirmacao({ ...confirmada, digital, [campo]: valor });
  }

  async function handleAnalisar() {
    if (await inativacao.analisar()) setEtapa(1);
  }

  async function handleExecutar() {
    const ok = await inativacao.executar(orfas > 0 && confirmada.orfas);
    if (ok) setEtapa(2);
  }

  function novaInativacao() {
    // Só existe depois de concluir, quando nada está em andamento: o reset nunca corta uma execução.
    inativacao.reset();
    setConfirmacao(SEM_CONFIRMACAO);
    setEtapa(0);
  }

  // Ao trocar de etapa o foco vai ao título: quem usa teclado ou leitor de tela começa a etapa do começo.
  const tituloRef = useRef<HTMLHeadingElement>(null);
  const primeiraRender = useRef(true);
  useEffect(() => {
    if (primeiraRender.current) {
      primeiraRender.current = false;
      return;
    }
    tituloRef.current?.focus({ preventScroll: true });
  }, [etapaAtual]);

  const detalhes = [
    inativacao.cadastro && inativacao.estruturas ? "2 bases enviadas" : "Cadastro e estruturas",
    resumo ? plural(resumo.executaveis, "a inativar", "a inativar") : "Conferir o impacto",
    concluido ? "ZIP gerado" : "Inativar e baixar",
  ];
  const passos: StepperItem[] = ETAPAS.map((e, i) => ({
    label: e.rotulo,
    detail: detalhes[i],
    state: concluido || i < etapaAtual ? "done" : i === etapaAtual ? "current" : "todo",
    selectable: !concluido && !travado && i < etapaAtual,
  }));

  const cabecalho = concluido
    ? { titulo: "Tudo pronto", descricao: "Comece outra inativação quando quiser." }
    : ETAPAS[etapaAtual];

  return (
    <div>
      <PageHeader
        title="Inativação"
        description="Inative usuários e mantenha as estruturas de aprovação consistentes."
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
        <ol className="list-decimal space-y-1 pl-5">
          <li>
            <strong>1º Passo:</strong> envie a <strong>base de cadastro</strong> e a <strong>base de estruturas</strong>{" "}
            (Excel) e cole <strong>CPFs</strong>, <strong>nomes completos</strong> ou <strong>e-mails</strong>, um por
            linha. Clique em <em>Analisar impacto</em>.
          </li>
          <li>
            <strong>2º Passo:</strong> confira, por usuário, a estrutura do viajante que será excluída e o efeito como
            aprovador (<em>compactação</em> ou <em>estrutura órfã</em>). Nada é alterado nesta etapa. Em caso de nomes
            repetidos, marque quem deve ser inativado.
          </li>
          <li>
            <strong>3º Passo:</strong> confirme e clique em <em>Executar inativação</em>. Você baixa um ZIP com a ficha de
            inativação e as estruturas atualizadas, prontos para a carga.
          </li>
        </ol>
        <p className="mt-2 text-text-subtle">
          Usuários sem CPF no cadastro não podem ser inativados por aqui: corrija o cadastro e analise de novo.
        </p>
      </Modal>

      <Stepper label="Etapas da inativação" steps={passos} onSelect={setEtapa} />

      <div className="grid items-start gap-6 xl:grid-cols-[minmax(0,1fr)_22rem]">
        <Card padding="lg">
          <header className="mb-4">
            <p className="text-xs font-medium text-text-muted">
              Etapa {etapaAtual + 1} de {ETAPAS.length}
            </p>
            <h3
              ref={tituloRef}
              id="inativacao_step_title"
              tabIndex={-1}
              className="text-lg font-semibold text-text outline-none"
            >
              {cabecalho.titulo}
            </h3>
            <p className="mt-0.5 text-sm text-text-muted">{cabecalho.descricao}</p>
          </header>

          {etapaAtual === 0 && (
            <>
              {/* O FileDropzone compartilhado não tem `disabled`: durante a análise ainda aceita arquivo, mas o
                  hook invalida a análise em curso e descarta a resposta antiga. */}
              <div className="grid gap-4 md:grid-cols-2">
                <FileDropzone
                  id="inativacao_cadastro"
                  containerId="inativacao_cadastro_uploadArea"
                  accept=".xlsx,.xls"
                  ariaLabel="Upload da base de cadastro. Pressione para selecionar arquivo"
                  description="Base de cadastro de usuários (.xlsx)"
                  syncFiles={inativacao.cadastro ? [inativacao.cadastro] : []}
                  onFiles={(list) => inativacao.setCadastro(list[0] ?? null)}
                >
                  <ArquivoEscolhido
                    file={inativacao.cadastro}
                    id="inativacao_cadastro"
                    disabled={travado}
                    onClear={() => inativacao.setCadastro(null)}
                  />
                </FileDropzone>
                <FileDropzone
                  id="inativacao_estruturas"
                  containerId="inativacao_estruturas_uploadArea"
                  accept=".xlsx,.xls"
                  ariaLabel="Upload da base de estruturas de aprovação. Pressione para selecionar arquivo"
                  description="Base de estruturas de aprovação (.xlsx)"
                  syncFiles={inativacao.estruturas ? [inativacao.estruturas] : []}
                  onFiles={(list) => inativacao.setEstruturas(list[0] ?? null)}
                >
                  <ArquivoEscolhido
                    file={inativacao.estruturas}
                    id="inativacao_estruturas"
                    disabled={travado}
                    onClear={() => inativacao.setEstruturas(null)}
                  />
                </FileDropzone>
              </div>

              <div className="mt-4">
                <Field id="lista_text" label="Quem será inativado (um por linha)">
                  <Textarea
                    id="lista_text"
                    rows={4}
                    placeholder={"Ex: João Silva\n12345678901\nusuario@example.com"}
                    aria-describedby="lista_valid_summary lista_duplicates_warning"
                    value={inativacao.listText}
                    disabled={travado}
                    onChange={(e) => inativacao.setListText(e.target.value)}
                  />
                </Field>
                <div id="lista_valid_summary" className="mt-2 text-sm text-text-muted">
                  <span className="rounded bg-accent/10 px-2 py-0.5 font-medium text-accent-text">
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
            </>
          )}

          {etapaAtual === 1 && analise && resumo && (
            <>
              <p id="inativacao_summary" className="mb-3 text-sm text-text">
                {plural(resumo.executaveis, "usuário será inativado", "usuários serão inativados")} ·{" "}
                {plural(resumo.estruturasExcluidas, "estrutura excluída", "estruturas excluídas")} ·{" "}
                {plural(resumo.estruturasCompactadas, "compactada", "compactadas")} ·{" "}
                <span className={orfas > 0 ? "font-medium text-danger" : ""}>
                  {plural(orfas, "estrutura órfã", "estruturas órfãs")}
                </span>
              </p>
              <ul id="inativacao_impacto" className="max-h-[28rem] space-y-3 overflow-y-auto pr-1">
                {analise.usuarios.map((u, i) => (
                  <ImpactoCard
                    key={`${u.cpf ?? u.nome}-${i}`}
                    usuario={u}
                    escolhidos={inativacao.escolhidos}
                    onEscolher={inativacao.escolher}
                    disabled={travado}
                  />
                ))}
              </ul>
              {pendentes > 0 && (
                <div className="mt-3 flex flex-wrap items-center gap-3 text-sm text-warning">
                  <span>
                    {plural(pendentes, "nome repetido aguarda", "nomes repetidos aguardam")} sua escolha. Para deixar um
                    de fora, volte e tire-o da lista.
                  </span>
                  <Button
                    id="inativacao_apply_selection_btn"
                    variant="outline"
                    size="sm"
                    loading={analisando}
                    disabled={inativacao.escolhidos.length === 0 || travado}
                    onClick={() => void inativacao.analisar()}
                  >
                    Aplicar seleção
                  </Button>
                </div>
              )}
            </>
          )}

          {etapaAtual === 2 && analise && resumo && !concluido && (
            <>
              <p className="text-sm text-text">
                Serão inativados <strong>{plural(resumo.executaveis, "usuário", "usuários")}</strong>, com{" "}
                {plural(resumo.estruturasExcluidas, "estrutura excluída", "estruturas excluídas")} e{" "}
                {plural(resumo.estruturasCompactadas, "compactada", "compactadas")}
                {orfas > 0 && (
                  <>
                    , e{" "}
                    <strong className="text-danger">
                      {plural(orfas, "estrutura sem aprovador", "estruturas sem aprovador")}
                    </strong>
                  </>
                )}
                .
              </p>
              <p className="mt-1 text-xs text-text-muted">
                Arquivos já carregados: {inativacao.cadastro?.name} e {inativacao.estruturas?.name}.
              </p>
              <div className="mt-4 space-y-2">
                <label className="flex items-start gap-2 text-sm text-text">
                  <input
                    id="inativacao_confirm_impacto"
                    type="checkbox"
                    className="mt-0.5 h-4 w-4"
                    checked={confirmada.impacto}
                    disabled={travado}
                    onChange={(e) => marcar("impacto", e.target.checked)}
                  />
                  <span>Revisei o impacto e quero inativar {plural(resumo.executaveis, "usuário", "usuários")}.</span>
                </label>
                {orfas > 0 && (
                  <label className="flex items-start gap-2 text-sm text-danger">
                    <input
                      id="inativacao_confirm_orfas"
                      type="checkbox"
                      className="mt-0.5 h-4 w-4"
                      checked={confirmada.orfas}
                      disabled={travado}
                      onChange={(e) => marcar("orfas", e.target.checked)}
                    />
                    <span>
                      Estou ciente de que {plural(orfas, "estrutura ficará", "estruturas ficarão")} sem nenhum aprovador
                      na Argo.
                    </span>
                  </label>
                )}
              </div>
              {executando && (
                <div className="mt-4">
                  <ProcessingProgress id="inativacao_progress" progress={inativacao.progress} />
                </div>
              )}
            </>
          )}

          {concluido && (
            <div
              id="inativacao_status"
              aria-live="polite"
              className="flex items-start gap-3 rounded-control border border-success/40 bg-success-soft px-4 py-3"
            >
              <span className="mt-0.5 text-success">
                <IconCheck className="h-5 w-5" />
              </span>
              <div className="text-sm">
                <p className="font-medium text-success">Inativação concluída</p>
                <p className="mt-0.5 text-text-muted">
                  O arquivo <strong className="font-medium text-text">{ARQUIVO_ZIP}</strong> foi gerado, com a ficha de
                  inativação e as estruturas atualizadas. Se o navegador não o salvou, baixe de novo em “Nesta sessão”.
                </p>
              </div>
            </div>
          )}

          {inativacao.failure && (
            <GenerationError
              id="inativacao_debug"
              className="mt-4"
              view={explainFailure(inativacao.failure, "Não foi possível concluir a inativação")}
            />
          )}

          <div className="mt-6 flex items-center justify-between gap-3 border-t border-border pt-4">
            {etapaAtual > 0 && !concluido ? (
              <Button id="inativacao_back_btn" variant="ghost" disabled={travado} onClick={() => setEtapa(etapaAtual - 1)}>
                Voltar
              </Button>
            ) : (
              <span />
            )}
            {concluido ? (
              <Button id="inativacao_new_btn" onClick={novaInativacao}>
                Nova inativação
              </Button>
            ) : etapaAtual === 0 ? (
              <Button
                id="inativacao_btn"
                aria-label="Analisar impacto da inativação"
                disabled={!inativacao.podeAnalisar || travado}
                loading={analisando}
                onClick={handleAnalisar}
              >
                {analisando ? "Analisando..." : "Analisar impacto"}
              </Button>
            ) : etapaAtual === 1 ? (
              <Button
                id="inativacao_next_btn"
                title={pendentes > 0 ? "Escolha quem inativar entre os nomes repetidos" : undefined}
                disabled={!podeContinuar(analise) || travado}
                onClick={() => setEtapa(2)}
              >
                Continuar
              </Button>
            ) : (
              <Button
                id="inativacao_execute_btn"
                variant={orfas > 0 ? "warning" : "primary"}
                disabled={!podeExecutar(analise, confirmada) || travado}
                loading={executando}
                onClick={handleExecutar}
              >
                {executando ? "Executando..." : "Executar inativação"}
              </Button>
            )}
          </div>
        </Card>

        <div className="xl:sticky xl:top-20">
          <RunHistoryPanel operation="inativacao" className="mt-0" />
        </div>
      </div>
    </div>
  );
}
