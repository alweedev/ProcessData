import type { GenerationFailure } from "../../lib/failure";
import { useRef, useState } from "react";
import { addLocalEntry } from "../../history/historyStore";
import {
  ApiRejection,
  postAnalysisSummary,
  postFormForBlob,
  type NameReviewItem,
  type QualityReport,
} from "../../lib/api";
import type { NameOverrides } from "../../lib/nameReview";
import { triggerAnchorDownload } from "../../lib/downloadFile";
import { addRun } from "../../runs/runsStore";
import { pushToast } from "../../toast/toastStore";

/** Nome do arquivo gerado: única fonte para o download, o registro da sessão e as mensagens da tela. */
export const OUTPUT_FILENAME = "saida_cadastro.xlsx";

export const MAX_FILES = 5;
const MAX_SIZE = 10 * 1024 * 1024;
const VALIDATE_TIMEOUT_MS = 15000;

/** O `accept` do <input> só vale para o seletor: arrastar e soltar aceita qualquer arquivo. */
export function isSpreadsheet(file: File): boolean {
  return /\.xlsx?$/i.test(file.name);
}

export function validateCadastroFiles(files: File[]): string | null {
  if (files.length > MAX_FILES) return `Máximo de ${MAX_FILES} arquivos por envio.`;
  for (const f of files) {
    if (!isSpreadsheet(f)) return `"${f.name}" não é uma planilha .xlsx ou .xls.`;
    if (f.size > MAX_SIZE) return `"${f.name}" excede 10 MB.`;
  }
  return null;
}

/** Aviso sobre arquivos que não são planilha, apontando o primeiro culpado pelo nome. */
function mensagemDeIgnorados(ignorados: File[], nenhumaValida: boolean): string {
  const [primeiro, ...resto] = ignorados;
  const alvo = resto.length
    ? `"${primeiro.name}" e mais ${resto.length} ${resto.length === 1 ? "arquivo" : "arquivos"}`
    : `"${primeiro.name}"`;
  const varios = ignorados.length > 1;
  return nenhumaValida
    ? `${alvo} ${varios ? "não são planilhas" : "não é uma planilha"} — só .xlsx ou .xls são aceitos.`
    : `${alvo} ${varios ? "foram ignorados" : "foi ignorado"} — só planilhas .xlsx ou .xls são aceitas.`;
}

interface ValidationState {
  status: "idle" | "loading" | "done" | "error";
  report: QualityReport | null;
  error: string | null;
  /** A planilha foi recusada pelo servidor (a geração falharia igual): o motivo, por arquivo. */
  failure: GenerationFailure | null;
  /** Nomes que o usuário precisa conferir antes de gerar (divisão duvidosa ou acima de 20 caracteres). */
  nameReview: NameReviewItem[];
}

const IDLE_VALIDATION: ValidationState = {
  status: "idle",
  report: null,
  error: null,
  failure: null,
  nameReview: [],
};

export function useCadastro() {
  const [files, setFiles] = useState<File[]>([]);
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [done, setDone] = useState(false);
  const [failure, setFailure] = useState<GenerationFailure | null>(null);
  const [validation, setValidation] = useState<ValidationState>(IDLE_VALIDATION);
  const abortRef = useRef<AbortController | null>(null);

  /** Arquivos que não são planilha são descartados um a um (com aviso pelo nome); o resto segue.
   *  Retorna `false` quando a seleção inteira é rejeitada (nenhuma planilha, arquivos demais /
   *  grandes demais): nada muda no que já estava escolhido — um arrasto errado não apaga o trabalho
   *  anterior — e o chamador deve refazer o <input> nativo, que ficou com a lista recusada. */
  function pickFiles(list: FileList): boolean {
    setFailure(null);
    const todos = Array.from(list);
    const candidates = todos.filter(isSpreadsheet);
    const ignorados = todos.filter((f) => !isSpreadsheet(f));
    if (candidates.length === 0) {
      pushToast(mensagemDeIgnorados(ignorados, true), "danger");
      return false;
    }
    const error = validateCadastroFiles(candidates);
    if (error) {
      pushToast(error, "danger");
      return false;
    }
    if (ignorados.length > 0) pushToast(mensagemDeIgnorados(ignorados, false), "info");
    setDone(false);
    invalidateValidation();
    setFiles(candidates);
    return true;
  }

  function removeFile(index: number) {
    setFailure(null);
    setFiles((prev) => prev.filter((_, i) => i !== index));
    setDone(false);
    setValidation(IDLE_VALIDATION);
  }

  /** Descarta o relatório e qualquer validação em andamento: ele vale só para as fichas,
   *  o login e o fluxo com que foi rodado. Sem soltar `abortRef`, a resposta atrasada de
   *  uma validação cancelada ainda reescreveria o estado. */
  function invalidateValidation() {
    abortRef.current?.abort();
    abortRef.current = null;
    setValidation(IDLE_VALIDATION);
  }

  function clear() {
    setFiles([]);
    setDone(false);
    setFailure(null);
    invalidateValidation();
  }

  /** Validação prévia (informativa): roda o pipeline no backend via
   *  `/api/analysis/summary` e mostra o relatório. NUNCA bloqueia a geração. */
  async function validate(loginChoice: string, fluxo: string) {
    if (!files.length) {
      pushToast("Selecione ao menos um arquivo para validar.", "danger");
      return;
    }
    abortRef.current?.abort();
    const controller = new AbortController();
    abortRef.current = controller;
    const timeoutId = window.setTimeout(() => controller.abort(), VALIDATE_TIMEOUT_MS);
    setValidation({ ...IDLE_VALIDATION, status: "loading" });
    try {
      const summary = await postAnalysisSummary(files, loginChoice, fluxo, controller.signal);
      if (abortRef.current !== controller) return; // invalidada ou substituída no meio do caminho
      setValidation({
        status: "done",
        report: summary.report,
        error: null,
        failure: null,
        nameReview: summary.name_review ?? [],
      });
    } catch (err) {
      if (abortRef.current !== controller) return;
      if (controller.signal.aborted) {
        setValidation({
          ...IDLE_VALIDATION,
          status: "error",
          error: "Validação cancelada ou expirada — a geração não depende disso.",
        });
      } else if (err instanceof ApiRejection) {
        // O servidor recusou a planilha: o motivo vai por arquivo (a tela o explica) e a geração falharia igual —
        // nada de "a geração não depende disso".
        setValidation({
          ...IDLE_VALIDATION,
          status: "error",
          failure: { message: err.message, fileErrors: err.fileErrors },
        });
      } else {
        const message = err instanceof Error ? err.message : String(err);
        setValidation({
          ...IDLE_VALIDATION,
          status: "error",
          error: `Não foi possível validar agora (${message}) — a geração não depende disso.`,
        });
        pushToast("Não foi possível validar a planilha agora.", "info");
      }
    } finally {
      window.clearTimeout(timeoutId);
    }
  }

  /** `nameOverrides`: a decisão do usuário na conferência de nomes — o backend a usa no lugar da sugestão e o
   *  vocabulário de nomes aprende com ela. */
  async function submit(loginChoice: string, fluxo: string, onSuccess?: () => void, nameOverrides?: NameOverrides) {
    if (!files.length) {
      pushToast("Selecione pelo menos um arquivo", "danger");
      return;
    }
    const error = validateCadastroFiles(files);
    if (error) {
      pushToast("Seleção de arquivos inválida (limite de arquivos ou tamanho excedido)", "danger");
      return;
    }
    setGenerating(true);
    setDone(false);
    setFailure(null);
    setProgress(0);
    const firstName = files[0].name;
    const count = files.length;
    try {
      const fd = new FormData();
      for (const f of files) fd.append("files[]", f);
      fd.append("login_choice", loginChoice);
      fd.append("fluxo", fluxo);
      if (nameOverrides && Object.keys(nameOverrides).length > 0) {
        fd.append("name_overrides", JSON.stringify(nameOverrides));
      }
      const blob = await postFormForBlob("/api/process_cadastro", fd, setProgress);
      const url = URL.createObjectURL(blob);
      triggerAnchorDownload(url, OUTPUT_FILENAME);
      setDone(true);
      setFiles([]);
      setValidation(IDLE_VALIDATION);
      addRun({
        operation: "cadastro",
        inputSummary: [count === 1 ? firstName : `${count} arquivos`, `Login ${loginChoice} · Fluxo ${fluxo}`],
        outputFilename: OUTPUT_FILENAME,
        blobUrl: url,
      });
      pushToast(`Cadastro concluído: login ${loginChoice}, fluxo ${fluxo}.`, "success");
      addLocalEntry(`Cadastro gerado: ${firstName} - ${new Date().toLocaleString("pt-BR")}`);
      onSuccess?.();
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setFailure({
        message,
        fileErrors: err instanceof ApiRejection ? err.fileErrors : undefined,
      });
      pushToast("Erro ao processar cadastro: " + message, "danger");
    } finally {
      setGenerating(false);
      setProgress(0);
    }
  }

  return {
    files,
    pickFiles,
    removeFile,
    clear,
    validate,
    validation,
    invalidateValidation,
    submit,
    generating,
    progress,
    done,
    failure,
  };
}
