export async function postFormJson<T>(url: string, formData: FormData): Promise<T> {
  const res = await fetch(url, { method: "POST", body: formData });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error || `Falha na requisição (${res.status})`);
  return data as T;
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
