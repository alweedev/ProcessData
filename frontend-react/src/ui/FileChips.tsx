import { humanFileSize } from "../lib/humanFileSize";

interface FileChipsProps {
  files: File[];
  onRemove: (index: number) => void;
  onClearAll?: () => void;
  clearAllId?: string;
  removeLabel?: (name: string) => string;
}

export function FileChips({ files, onRemove, onClearAll, clearAllId, removeLabel }: FileChipsProps) {
  if (files.length === 0) return null;
  return (
    <div className="mt-3 flex flex-wrap items-center gap-2">
      {files.map((file, index) => (
        <span
          key={`${file.name}-${index}`}
          className="inline-flex items-center gap-2 rounded-pill border border-border bg-surface-2 py-1 pl-3 pr-1.5 text-sm text-text"
        >
          <span className="max-w-[16rem] truncate">{file.name}</span>
          <span className="text-xs text-text-subtle">{humanFileSize(file.size)}</span>
          <button
            type="button"
            aria-label={removeLabel ? removeLabel(file.name) : `Remover ${file.name}`}
            onClick={(e) => {
              e.stopPropagation();
              onRemove(index);
            }}
            className="flex h-5 w-5 items-center justify-center rounded-pill text-text-subtle transition-colors hover:bg-danger/10 hover:text-danger"
          >
            ✕
          </button>
        </span>
      ))}
      {onClearAll && (
        <button
          type="button"
          id={clearAllId}
          aria-label="Remover todos os arquivos"
          onClick={(e) => {
            e.stopPropagation();
            onClearAll();
          }}
          className="text-xs font-medium text-text-muted underline-offset-2 hover:text-danger hover:underline"
        >
          Limpar todos
        </button>
      )}
    </div>
  );
}
