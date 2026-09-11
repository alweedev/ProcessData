import { useRef, useState } from "react";
import { postAnalysisSummary, postFormForBlob, type QualityReport } from "../../lib/api";
import { addRun } from "../../runs/runsStore";
import { pushToast } from "../../toast/toastStore";

const MAX_FILES = 5;
const MAX_SIZE = 10 * 1024 * 1024;
const VALIDATE_TIMEOUT_MS = 15000;

export function validateCadastroFiles(files: File[]): string | null {
  if (files.length > MAX_FILES) return `Máximo de ${MAX_FILES} arquivos por envio.`;
  for (const f of files) {
    if (f.size > MAX_SIZE) return `"${f.name}" excede 10 MB.`;
  }
  return null;
}

interface ValidationState {
  status: "idle" | "loading" | "done" | "error";
  report: QualityReport | null;
  error: string | null;
}

const IDLE_VALIDATION: ValidationState = { status: "idle", report: null, error: null };

export function useCadastro() {
  const [files, setFiles] = useState<File[]>([]);
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [done, setDone] = useState(false);
  const [debugMsg, setDebugMsg] = useState("");
  const [validation, setValidation] = useState<ValidationState>(IDLE_VALIDATION);
  const abortRef = useRef<AbortController | null>(null);

  /** Retorna `false` quando a seleção é rejeitada (arquivos demais / grandes
   *  demais) — o chamador deve então forçar a limpeza do <input> nativo,
   *  já que setFiles([]) sozinho não reflete no FileList real do browser. */
  function pickFiles(list: FileList): boolean {
    const candidates = Array.from(list);
    const error = validateCadastroFiles(candidates);
    if (error) {
      pushToast(error, "danger");
      setFiles([]);
      setValidation(IDLE_VALIDATION);
      return false;
    }
    setDone(false);
    setValidation(IDLE_VALIDATION);
    setFiles(candidates);
    return true;
  }

  function removeFile(index: number) {
    setFiles((prev) => prev.filter((_, i) => i !== index));
    setDone(false);
    setValidation(IDLE_VALIDATION);
  }

  function clear() {
    setFiles([]);
    setDone(false);
    setDebugMsg("");
    setValidation(IDLE_VALIDATION);
    abortRef.current?.abort();
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
    setValidation({ status: "loading", report: null, error: null });
    try {
      const summary = await postAnalysisSummary(files, loginChoice, fluxo, controller.signal);
      setValidation({ status: "done", report: summary.report, error: null });
    } catch (err) {
      if (controller.signal.aborted) {
        setValidation({
          status: "error",
          report: null,
          error: "Validação cancelada ou expirada — a geração não depende disso.",
        });
      } else {
        const message = err instanceof Error ? err.message : String(err);
        setValidation({
          status: "error",
          report: null,
          error: `Não foi possível validar agora (${message}) — a geração não depende disso.`,
        });
        pushToast("Não foi possível validar a planilha agora.", "info");
      }
    } finally {
      window.clearTimeout(timeoutId);
    }
  }

  async function submit(loginChoice: string, fluxo: string, onSuccess?: () => void) {
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
    setDebugMsg("");
    setProgress(0);
    const firstName = files[0].name;
    const count = files.length;
    try {
      const fd = new FormData();
      for (const f of files) fd.append("files[]", f);
      fd.append("login_choice", loginChoice);
      fd.append("fluxo", fluxo);
      const blob = await postFormForBlob("/api/process_cadastro", fd, setProgress);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "saida_cadastro.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setDone(true);
      setFiles([]);
      setValidation(IDLE_VALIDATION);
      addRun({
        operation: "cadastro",
        inputSummary: [count === 1 ? firstName : `${count} arquivos`, `Login ${loginChoice} · Fluxo ${fluxo}`],
        outputFilename: "saida_cadastro.xlsx",
        blobUrl: url,
      });
      pushToast(`Cadastro processado! Opções: ${loginChoice}, ${fluxo}.`, "success");
      window.addToHistory?.(`Cadastro gerado: ${firstName} - ${new Date().toLocaleString("pt-BR")}`);
      onSuccess?.();
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setDebugMsg("Erro: " + message);
      pushToast("Erro ao processar cadastro: " + message, "danger");
    } finally {
      setGenerating(false);
      setProgress(0);
    }
  }

  return { files, pickFiles, removeFile, clear, validate, validation, submit, generating, progress, done, debugMsg };
}
