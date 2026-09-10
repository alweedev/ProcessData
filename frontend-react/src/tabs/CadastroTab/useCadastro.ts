import { useState } from "react";
import { postFormForBlob } from "../../lib/api";
import { pushToast } from "../../toast/toastStore";

const MAX_FILES = 5;
const MAX_SIZE = 10 * 1024 * 1024;

export function validateCadastroFiles(files: File[]): string | null {
  if (files.length > MAX_FILES) return `Máximo de ${MAX_FILES} arquivos por envio.`;
  for (const f of files) {
    if (f.size > MAX_SIZE) return `"${f.name}" excede 10 MB.`;
  }
  return null;
}

export function useCadastro() {
  const [files, setFiles] = useState<File[]>([]);
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [done, setDone] = useState(false);
  const [debugMsg, setDebugMsg] = useState("");

  /** Retorna `false` quando a seleção é rejeitada (arquivos demais / grandes
   *  demais) — o chamador deve então forçar a limpeza do <input> nativo,
   *  já que setFiles([]) sozinho não reflete no FileList real do browser. */
  function pickFiles(list: FileList): boolean {
    const candidates = Array.from(list);
    const error = validateCadastroFiles(candidates);
    if (error) {
      pushToast(error, "danger");
      setFiles([]);
      return false;
    }
    setDone(false);
    setFiles(candidates);
    return true;
  }

  function clear() {
    setFiles([]);
    setDone(false);
    setDebugMsg("");
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
      URL.revokeObjectURL(url);
      setDone(true);
      setFiles([]);
      pushToast(`Cadastro processado! Opções: ${loginChoice}, ${fluxo}.`, "success");
      window.addToHistory?.(`Cadastro gerado: ${files[0].name} - ${new Date().toLocaleString("pt-BR")}`);
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

  return { files, pickFiles, clear, submit, generating, progress, done, debugMsg };
}
