import { useMemo, useRef, useState } from "react";
import { triggerAnchorDownload } from "../../lib/downloadFile";
import type { GenerationFailure } from "../../lib/failure";
import { InativacaoApiError, postAnalisar, postExecutar, type AnaliseInativacao } from "../../lib/inativacaoApi";
import { addRun } from "../../runs/runsStore";
import { pushToast } from "../../toast/toastStore";

export const ARQUIVO_ZIP = "inativacao.zip";

export interface Classification {
  validCpfs: string[];
  validNames: string[];
  validEmails: string[];
  duplicates: string[];
  /** Linhas que não são CPF (11 dígitos), e-mail nem nome completo: o servidor as ignora, a tela avisa. */
  invalid: string[];
  totalValid: number;
}

const EMAIL_PATTERN = /^[^@\s]+@[^@\s]+\.[^@\s]+$/i;

export function isValidFullName(s: string): boolean {
  const norm = s.normalize("NFKD").replace(/\p{Diacritic}/gu, "");
  const parts = norm.trim().split(/\s+/);
  return parts.length >= 2 && norm.trim().length >= 3;
}

export function classifyList(text: string): Classification {
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
  const invalid = nameRaw.filter((l) => !isValidFullName(l));
  return { validCpfs, validNames, validEmails: emailRaw, duplicates, invalid, totalValid: validCpfs.length + validNames.length + emailRaw.length };
}

export function formatCpf(c: string): string {
  return c && c.length === 11 ? `${c.slice(0, 3)}.${c.slice(3, 6)}.${c.slice(6, 9)}-${c.slice(9)}` : c || "";
}

function comoFalha(err: unknown): GenerationFailure {
  return { message: err instanceof Error ? err.message : String(err) };
}

export function useInativacao() {
  const [listText, setListTextRaw] = useState("");
  const [cadastro, setCadastroRaw] = useState<File | null>(null);
  const [estruturas, setEstruturasRaw] = useState<File | null>(null);
  const [analise, setAnalise] = useState<AnaliseInativacao | null>(null);
  const [escolhidos, setEscolhidos] = useState<string[]>([]);
  // Nome da coluna que o servidor detectou como candidata a CPF do viajante (ex.: "Valor"), pendente de confirmação.
  const [cpfViajanteCandidato, setCpfViajanteCandidato] = useState<string | null>(null);
  const [analisando, setAnalisando] = useState(false);
  const [executando, setExecutando] = useState(false);
  const [progress, setProgress] = useState(0);
  const [failure, setFailure] = useState<GenerationFailure | null>(null);
  const [concluido, setConcluido] = useState(false);
  // Geração da entrada: sobe a cada invalidar()/reset(). Resposta de requisição de geração antiga é descartada.
  const geracao = useRef(0);
  // Refs (não estado) para a trava de reentrada valer mesmo em dois cliques antes de re-renderizar.
  const analisandoRef = useRef(false);
  const executandoRef = useRef(false);
  // Confirmação de CPF do viajante usada na última análise bem-sucedida: `executar()` reenvia o mesmo valor.
  const confirmarCpfViajanteRef = useRef(false);

  const classification = useMemo(() => classifyList(listText), [listText]);
  const itens = useMemo(
    () => [...classification.validCpfs, ...classification.validNames, ...classification.validEmails],
    [classification],
  );
  const podeAnalisar = Boolean(cadastro) && Boolean(estruturas) && itens.length > 0 && !analisando;

  /** Mexer nas entradas invalida a análise: o que o operador viu deixa de valer. */
  function invalidar() {
    geracao.current += 1;
    analisandoRef.current = false;
    setAnalisando(false);
    setAnalise(null);
    setEscolhidos([]);
    setFailure(null);
    setCpfViajanteCandidato(null);
    confirmarCpfViajanteRef.current = false;
  }

  function setListText(value: string) {
    setListTextRaw(value);
    invalidar();
  }

  function setCadastro(file: File | null) {
    setCadastroRaw(file);
    invalidar();
  }

  function setEstruturas(file: File | null) {
    setEstruturasRaw(file);
    invalidar();
  }

  function escolher(cpf: string, marcado: boolean) {
    setEscolhidos((atuais) => (marcado ? [...new Set([...atuais, cpf])] : atuais.filter((c) => c !== cpf)));
  }

  async function analisar(confirmarCpfViajante = confirmarCpfViajanteRef.current): Promise<boolean> {
    if (!cadastro || !estruturas || analisandoRef.current) return false;
    const minha = geracao.current;
    analisandoRef.current = true;
    setAnalisando(true);
    setFailure(null);
    try {
      const resultado = await postAnalisar(cadastro, estruturas, itens, escolhidos, confirmarCpfViajante);
      // As entradas mudaram durante a requisição: o resultado é de outra lista e não vale.
      if (minha !== geracao.current) return false;
      setAnalise(resultado);
      setCpfViajanteCandidato(null);
      confirmarCpfViajanteRef.current = confirmarCpfViajante;
      return true;
    } catch (err) {
      if (minha !== geracao.current) return false;
      setAnalise(null);
      if (err instanceof InativacaoApiError && err.code === "CPF_VIAJANTE_A_CONFIRMAR") {
        setCpfViajanteCandidato((err.extra?.colunaCandidata as string) || "Valor");
        return false;
      }
      setFailure(comoFalha(err));
      return false;
    } finally {
      // Só quem ainda é a última chamada limpa a flag (invalidar()/reset() já a limparam para as antigas).
      if (minha === geracao.current) {
        analisandoRef.current = false;
        setAnalisando(false);
      }
    }
  }

  /** Reanalisa autorizando o uso da coluna candidata (ex.: "Valor") como CPF do viajante. */
  async function confirmarCpfViajanteEAnalisar(): Promise<boolean> {
    return analisar(true);
  }

  async function executar(ignorarOrfas: boolean): Promise<boolean> {
    if (!cadastro || !estruturas || !analise || executandoRef.current) return false;
    const minha = geracao.current;
    executandoRef.current = true;
    const cpfs = analise.usuarios
      .filter((u) => u.situacao === "EXECUTAVEL" && u.cpf)
      .map((u) => u.cpf as string);
    setExecutando(true);
    setProgress(0);
    setFailure(null);
    try {
      const blob = await postExecutar(
        cadastro,
        estruturas,
        cpfs,
        analise.impressaoDigital,
        ignorarOrfas,
        confirmarCpfViajanteRef.current,
        setProgress,
      );
      const url = URL.createObjectURL(blob);
      triggerAnchorDownload(url, ARQUIVO_ZIP);
      addRun({
        operation: "inativacao",
        inputSummary: [cadastro.name, estruturas.name, `${cpfs.length} usuário(s)`],
        outputFilename: ARQUIVO_ZIP,
        blobUrl: url,
      });
      pushToast("Inativação executada.", "success");
      setConcluido(true);
      return true;
    } catch (err) {
      // Entradas mudaram durante a execução: a falha é de um estado que o operador já descartou.
      if (minha !== geracao.current) return false;
      // A análise que o operador viu não vale mais: volta a exigir uma nova.
      if (err instanceof InativacaoApiError && err.code === "ANALISE_DIVERGENTE") setAnalise(null);
      setFailure(comoFalha(err));
      return false;
    } finally {
      executandoRef.current = false;
      setExecutando(false);
      setProgress(0);
    }
  }

  function reset() {
    geracao.current += 1;
    analisandoRef.current = false;
    setAnalisando(false);
    setListTextRaw("");
    setCadastroRaw(null);
    setEstruturasRaw(null);
    setAnalise(null);
    setEscolhidos([]);
    setFailure(null);
    setConcluido(false);
    setCpfViajanteCandidato(null);
    confirmarCpfViajanteRef.current = false;
  }

  return {
    listText,
    setListText,
    cadastro,
    setCadastro,
    estruturas,
    setEstruturas,
    classification,
    itens,
    podeAnalisar,
    analise,
    escolhidos,
    escolher,
    analisando,
    executando,
    progress,
    failure,
    concluido,
    cpfViajanteCandidato,
    analisar,
    confirmarCpfViajanteEAnalisar,
    executar,
    reset,
  };
}
