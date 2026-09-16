/** Dispara o download de uma URL (blob ou não) via link temporário. Não
 * revoga a URL -- use quando o chamador retém o blob (ex.: "baixar de novo"
 * em runsStore); para um blob de uso único, prefira `downloadFile`. */
export function triggerAnchorDownload(url: string, filename: string) {
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

export function downloadFile(content: BlobPart, filename: string, mimeType: string) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  triggerAnchorDownload(url, filename);
  URL.revokeObjectURL(url);
}
