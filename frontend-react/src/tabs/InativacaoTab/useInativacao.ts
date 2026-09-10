import { useMemo, useState } from "react";
import { postFormForBlob, postFormJson } from "../../lib/api";
import { pushToast } from "../../toast/toastStore";

export interface InativacaoResult {
  id: string | null;
  nome: string;
  cpf: string;
  email: string;
  status_atual: string;
  found: boolean;
}

interface Classification {
  validCpfs: string[];
  validNames: string[];
  validEmails: string[];
  duplicates: string[];
  totalValid: number;
}

const EMAIL_PATTERN = /^[^@\s]+@[^@\s]+\.[^@\s]+$/i;

function isValidFullName(s: string): boolean {
  const norm = s.normalize("NFKD").replace(/\p{Diacritic}/gu, "");
  const parts = norm.trim().split(/\s+/);
  return parts.length >= 2 && norm.trim().length >= 3;
}

function classifyList(text: string): Classification {
  const lines = (text || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const cpfRaw: string[] = [];
  const nameRaw: string[] = [];
  const emailRaw: string[] = [];
  for (const line of lines) {
    const digits = line.replace(/\D/g, "");
    if (EMAIL_PATTERN.test(line)) emailRaw.push(line);
    else if (digits.length === 11) cpfRaw.push(digits);
    else nameRaw.push(line);
  }

  const seen = new Set<string>();
  const duplicates: string[] = [];
  const validCpfs: string[] = [];
  for (const c of cpfRaw) {
    if (seen.has(c)) {
      if (!duplicates.includes(c)) duplicates.push(c);
    } else {
      seen.add(c);
      validCpfs.push(c);
    }
  }

  const validNames = nameRaw.filter(isValidFullName);
  return { validCpfs, validNames, validEmails: emailRaw, duplicates, totalValid: validCpfs.length + validNames.length + emailRaw.length };
}

export function formatCpf(c: string): string {
  return c && c.length === 11 ? `${c.slice(0, 3)}.${c.slice(3, 6)}.${c.slice(6, 9)}-${c.slice(9)}` : c || "";
}

export function useInativacao() {
  const [listText, setListText] = useState("");
  const [baseFile, setBaseFile] = useState<File | null>(null);
  const [results, setResults] = useState<InativacaoResult[]>([]);
  const [searching, setSearching] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [debugMsg, setDebugMsg] = useState("");

  const classification = useMemo(() => classifyList(listText), [listText]);
  const canSearch = Boolean(baseFile) && classification.totalValid > 0;

  async function search() {
    if (!baseFile) {
      pushToast("Envie a base.", "danger");
      return;
    }
    if (!classification.totalValid) {
      pushToast("Nenhum CPF, Nome ou E-mail válido para buscar.", "danger");
      return;
    }
    setSearching(true);
    setDebugMsg("");
    try {
      const fd = new FormData();
      fd.append("base", baseFile);
      fd.append(
        "itens",
        JSON.stringify([...classification.validCpfs, ...classification.validNames, ...classification.validEmails]),
      );
      const data = await postFormJson<{ items?: InativacaoResult[] }>("/api/inativacao/buscar", fd);
      const items = Array.isArray(data.items) ? data.items : [];
      setResults(items);
      pushToast(`Busca concluída: ${items.length} itens.`, "success");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setDebugMsg("Erro: " + message);
      pushToast("Erro na busca: " + message, "danger");
    } finally {
      setSearching(false);
    }
  }

  async function generate(onDone?: () => void) {
    if (!baseFile) {
      pushToast("Envie a base para gerar a inativação.", "danger");
      return;
    }
    const finalListText = listText || [...classification.validCpfs, ...classification.validNames].join("\n");
    setGenerating(true);
    setProgress(0);
    try {
      const fd = new FormData();
      fd.append("base", baseFile);
      fd.append("lista_text", finalListText);
      const blob = await postFormForBlob("/api/process_inativacao", fd, setProgress);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "saida_inativacao.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      pushToast("Inativação processada.", "success");
      onDone?.();
    } catch (err) {
      pushToast(err instanceof Error ? err.message : String(err), "danger");
    } finally {
      setGenerating(false);
      setProgress(0);
    }
  }

  return {
    listText,
    setListText,
    baseFile,
    setBaseFile,
    classification,
    canSearch,
    results,
    setResults,
    searching,
    generating,
    progress,
    debugMsg,
    search,
    generate,
  };
}
