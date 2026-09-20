import { humanFileSize } from "../lib/humanFileSize";
import { IconChip } from "./IconChip";
import { IconFile, IconX } from "./icons";

interface FileListProps {
  files: File[];
  maxFiles: number;
  onRemove: (index: number) => void;
  onClearAll?: () => void;
  clearAllId?: string;
}

/** Arquivos escolhidos, um por linha (nome, tamanho e remover), com resumo
 *  de quantidade/tamanho total — substitui os chips soltos (FileChips). */
export function FileList({ files, maxFiles, onRemove, onClearAll, clearAllId }: FileListProps) {
  if (files.length === 0) return null;
  const total = files.reduce((sum, file) => sum + file.size, 0);

  return (
    <div className="mt-4">
      <div className="mb-2 flex items-center justify-between gap-3 text-xs text-text-muted">
        <span>
          <strong className="font-semibold text-text">
            {files.length} de {maxFiles}
          </strong>{" "}
          {files.length === 1 ? "arquivo" : "arquivos"} · {humanFileSize(total)} no total
        </span>
        {onClearAll && (
          <button
            type="button"
            id={clearAllId}
            aria-label="Remover todos os arquivos"
            onClick={onClearAll}
            className="font-medium underline-offset-2 hover:text-danger hover:underline"
          >
            Limpar todos
          </button>
        )}
      </div>
      <ul className="divide-y divide-border overflow-hidden rounded-control border border-border">
        {files.map((file, index) => (
          <li key={`${file.name}-${index}`} className="flex items-center gap-3 bg-surface px-3 py-2">
            <IconChip icon={<IconFile className="h-4 w-4" />} size="sm" tone="success" />
            <span className="min-w-0 flex-1 truncate text-sm text-text" title={file.name}>
              {file.name}
            </span>
            <span className="shrink-0 text-xs tabular-nums text-text-subtle">{humanFileSize(file.size)}</span>
            <button
              type="button"
              aria-label={`Remover ${file.name}`}
              onClick={() => onRemove(index)}
              className="flex h-7 w-7 shrink-0 items-center justify-center rounded-control text-text-subtle outline-none transition-colors hover:bg-danger/10 hover:text-danger focus-visible:ring-2 focus-visible:ring-accent/50"
            >
              <IconX className="h-4 w-4" />
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
