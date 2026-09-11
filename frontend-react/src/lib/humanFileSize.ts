const UNITS = ["B", "KB", "MB", "GB"];

/** Tamanho legível: 0 B, 640 KB, 2.4 MB, ... */
export function humanFileSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const i = Math.min(UNITS.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const value = bytes / 1024 ** i;
  const rounded = value >= 10 || i === 0 ? Math.round(value).toString() : value.toFixed(1);
  return `${rounded} ${UNITS[i]}`;
}
