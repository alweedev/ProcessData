import { ApiRejection, postFormForBlob } from "./api";

export type Situacao = "EXECUTAVEL" | "SEM_CPF" | "JA_INATIVO" | "NAO_LOCALIZADO" | "PENDENTE_SELECAO";

export interface Candidato {
  cpf: string;
  cpfMascarado: string;
  nome: string;
  email: string;
}

export interface EstruturaComoAprovador {
  aprovacaoId: string;
  posicoes: number[];
  segundoNivel: boolean;
  acao: "COMPACTACAO" | "ORFA";
}

export interface UsuarioAnalise {
  cpf: string | null;
  cpfMascarado: string;
  nome: string;
  email: string;
  situacao: Situacao;
  alerta: string | null;
  estruturasViajante: string[];
  comoAprovador: EstruturaComoAprovador[];
  candidatos: Candidato[];
}

export interface ResumoAnalise {
  executaveis: number;
  estruturasExcluidas: number;
  estruturasCompactadas: number;
  estruturasOrfas: number;
  duplicados: string[];
}

export interface AnaliseInativacao {
  usuarios: UsuarioAnalise[];
  resumo: ResumoAnalise;
  impressaoDigital: string;
}

/** Erro de negócio da inativação: a tela decide pelo `code` estável do servidor, não pelo texto. */
export class InativacaoApiError extends Error {
  readonly code: string;

  constructor(message: string, code: string) {
    super(message);
    this.name = "InativacaoApiError";
    this.code = code;
  }
}

const MENSAGEM_GENERICA = "Não foi possível concluir a operação. Verifique a conexão e tente de novo.";

function basesForm(cadastro: File, estruturas: File): FormData {
  const fd = new FormData();
  fd.append("cadastro", cadastro);
  fd.append("estruturas", estruturas);
  return fd;
}

/** Diagnóstico de impacto: não altera nada no servidor. */
export async function postAnalisar(
  cadastro: File,
  estruturas: File,
  itens: string[],
  selecionados: string[],
): Promise<AnaliseInativacao> {
  const fd = basesForm(cadastro, estruturas);
  fd.append("itens", JSON.stringify(itens));
  fd.append("selecionados", JSON.stringify(selecionados));
  const res = await fetch("/api/inativacao/analisar", { method: "POST", body: fd });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new InativacaoApiError(data?.error || `Falha na requisição (${res.status})`, data?.code || "ERRO_INTERNO");
  }
  if (!Array.isArray(data?.usuarios)) throw new InativacaoApiError("Resposta inesperada do servidor.", "ERRO_INTERNO");
  return data as AnaliseInativacao;
}

/** Efetiva a inativação: devolve o ZIP com a ficha e as estruturas atualizadas. */
export async function postExecutar(
  cadastro: File,
  estruturas: File,
  cpfs: string[],
  impressaoDigital: string,
  ignorarOrfas: boolean,
  onProgress?: (pct: number) => void,
): Promise<Blob> {
  const fd = basesForm(cadastro, estruturas);
  fd.append("cpfs", JSON.stringify(cpfs));
  fd.append("impressaoDigital", impressaoDigital);
  fd.append("ignore_orphan_warning", ignorarOrfas ? "true" : "false");
  try {
    return await postFormForBlob("/api/inativacao/executar", fd, onProgress);
  } catch (err) {
    if (err instanceof ApiRejection) throw new InativacaoApiError(err.message, err.code ?? "ERRO_INTERNO");
    // Corpo não-JSON (proxy/HTML), falha de rede ou arquivo inválido: nunca repassa o texto cru ao usuário.
    throw new InativacaoApiError(MENSAGEM_GENERICA, "ERRO_INTERNO");
  }
}
