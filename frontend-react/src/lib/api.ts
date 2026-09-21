export async function postFormJson<T>(url: string, formData: FormData): Promise<T> {
  const res = await fetch(url, { method: "POST", body: formData });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error || `Falha na requisição (${res.status})`);
  return data as T;
}

/** Uma linha com problema: onde está, quem é o passageiro e cada problema em separado. */
export interface LineDetail {
  /** "Linha 4" (com vários arquivos: "Arquivo 2 · linha 4") */
  label: string;
  /** nome do passageiro como no documento (vazio se a linha não tinha nome) */
  nome: string;
  erros: string[];
  /** a linha não tem nenhum campo obrigatório preenchido (só o nome): parece não preenchida. Ausente em servidor antigo. */
  sem_preenchimento?: boolean;
}

/** Relatório de qualidade da planilha — shape idêntico a
 *  backend/services/report_service.py::build_quality_report. */
export interface QualityReport {
  total_rows: number;
  valid_rows: number;
  invalid_rows: number;
  /** linhas repetidas (mesmo Login + Nome completo) que o backend já tirou do arquivo de carga */
  duplicated_rows: number;
  /** string única já concatenada (não é lista) */
  general_errors: string;
  /** { "Linha 4": "msg; msg" } — a linha do Excel (com vários arquivos: "Arquivo 2 · linha 4") */
  line_errors: Record<string, string>;
  /** O mesmo por linha, com o passageiro e cada problema separado (ausente em servidor antigo: ver `lineDetails`). */
  line_details?: LineDetail[];
  /** { "<campo obrigatório, como na ficha>": <qtd de células em branco> }, ex.: "Telefone": 3 */
  required_blank: Record<string, number>;
}

/** Um nome que o usuário precisa conferir antes de gerar: a divisão entre Nome e Sobrenome é duvidosa, ou um dos
 *  dois passa de 20 caracteres (o backend não corta: quem confere ajusta). */
export interface NameReviewItem {
  /** "arquivo:linha" — identifica a linha da ficha; volta em `name_overrides` na geração */
  key: string;
  label: string;
  /** nome como no documento (já limpo: maiúsculo, sem acento) */
  nome_completo: string;
  nome: string;
  sobrenome: string;
  confianca: "alta" | "baixa";
  motivos: string[];
  estouro: { nome: boolean; sobrenome: boolean };
  /** proposta de abreviação quando passou de 20 (nomes do meio viram inicial) */
  sugestao: { nome: string; sobrenome: string } | null;
}

export interface AnalysisSummary {
  report: QualityReport;
  /** até 20 linhas do resultado processado (todas as colunas do modelo) */
  preview: Record<string, unknown>[];
  name_review: NameReviewItem[];
}

/** O servidor respondeu e RECUSOU a planilha (vazia, quebrada, parâmetro inválido...), com uma mensagem já
 *  pensada para o usuário. Diferente de falha de rede/timeout: aqui a geração falharia pelo mesmo motivo. */
export class ApiRejection extends Error {
  /** `{nome do arquivo: motivo}`, quando o servidor separa o problema por arquivo. */
  readonly fileErrors?: Record<string, string>;
  /** Código estável do erro de negócio (`code`), quando o servidor o envia: a tela decide por ele, não pelo texto. */
  readonly code?: string;

  constructor(message: string, fileErrors?: Record<string, string>, code?: string) {
    super(message);
    this.fileErrors = fileErrors;
    this.code = code;
  }
}

function asFileErrors(value: unknown): Record<string, string> | undefined {
  if (!value || typeof value !== "object") return undefined;
  const entries = Object.entries(value).filter((entry): entry is [string, string] => typeof entry[1] === "string");
  return entries.length > 0 ? Object.fromEntries(entries) : undefined;
}

/**
 * Validação prévia da planilha de cadastro: roda o mesmo pipeline de
 * `/api/process_cadastro` mas devolve só o relatório (nenhum arquivo é gerado).
 * Reusa `POST /api/analysis/summary`, que já existe no backend.
 */
export async function postAnalysisSummary(
  files: File[],
  loginChoice: string,
  fluxo: string,
  signal?: AbortSignal,
): Promise<AnalysisSummary> {
  const fd = new FormData();
  for (const f of files) fd.append("files[]", f);
  fd.append("login_choice", loginChoice);
  fd.append("fluxo", fluxo);
  const res = await fetch("/api/analysis/summary", {
    method: "POST",
    body: fd,
    signal,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new ApiRejection(data?.error || `Falha ao validar (${res.status})`, asFileErrors(data?.errors));
  return data as AnalysisSummary;
}

/**
 * Upload multipart com progresso real (via XHR, `fetch` não expõe upload
 * progress de forma confiável) que espera de volta um arquivo binário
 * (blob), não JSON — usado pelos endpoints de geração/exportação.
 */
export function postFormForBlob(url: string, formData: FormData, onProgress?: (pct: number) => void): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.upload.addEventListener("progress", (e) => {
      if (e.lengthComputable && onProgress) onProgress((e.loaded / e.total) * 100);
    });
    // Arquivos pequenos podem terminar sem nenhum evento de progresso; o "load"
    // do upload garante o 100% (a resposta do servidor vem depois, em onload).
    xhr.upload.addEventListener("load", () => onProgress?.(100));
    xhr.open("POST", url, true);
    xhr.responseType = "blob";
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        const blob = xhr.response as Blob;
        if (!blob || blob.size === 0) {
          reject(new Error("Arquivo gerado inválido."));
          return;
        }
        resolve(blob);
        return;
      }
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const obj = JSON.parse(String(reader.result || "{}"));
          reject(
            new ApiRejection(
              obj.error || `Erro ${xhr.status}`,
              asFileErrors(obj.errors),
              typeof obj.code === "string" ? obj.code : undefined,
            ),
          );
        } catch {
          reject(new Error(String(reader.result || `Erro ${xhr.status}`)));
        }
      };
      reader.onerror = () => reject(new Error(`Erro ${xhr.status}`));
      reader.readAsText(xhr.response);
    };
    xhr.onerror = () => reject(new Error("Erro de rede."));
    xhr.send(formData);
  });
}
