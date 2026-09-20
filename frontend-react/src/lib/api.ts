export async function postFormJson<T>(url: string, formData: FormData): Promise<T> {
  const res = await fetch(url, { method: "POST", body: formData });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error || `Falha na requisição (${res.status})`);
  return data as T;
}

/** Relatório de qualidade da planilha — shape idêntico a
 *  backend/services/report_service.py::build_quality_report. */
export interface QualityReport {
  total_rows: number;
  valid_rows: number;
  invalid_rows: number;
  duplicated_rows: number;
  /** string única já concatenada (não é lista) */
  general_errors: string;
  /** { "<índice da linha>": "msg; msg" } */
  line_errors: Record<string, string>;
  /** { "<Coluna>": <qtd de células em branco> } */
  required_blank: Record<string, number>;
}

export interface AnalysisSummary {
  report: QualityReport;
  /** até 20 linhas do resultado processado (todas as colunas do modelo) */
  preview: Record<string, unknown>[];
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
  const res = await fetch("/api/analysis/summary", { method: "POST", body: fd, signal });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error || `Falha ao validar (${res.status})`);
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
          reject(new Error(obj.error || `Erro ${xhr.status}`));
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
