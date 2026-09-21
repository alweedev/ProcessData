import { useEffect, useState, type CSSProperties } from "react";
import { humanFileSize } from "../lib/humanFileSize";
import { cn } from "./cn";
import { IconChip } from "./IconChip";
import { IconFile, IconX } from "./icons";

interface FileListProps {
  files: File[];
  maxFiles: number;
  onRemove: (index: number) => void;
  onClearAll?: () => void;
  clearAllId?: string;
}

const STAGGER_MS = 80;

// Chave estável por arquivo: com `nome-índice`, remover a 1ª linha remontava as demais.
const ids = new WeakMap<File, number>();
let nextId = 0;
function idOf(file: File): number {
  let id = ids.get(file);
  if (id === undefined) {
    id = nextId++;
    ids.set(file, id);
  }
  return id;
}

interface FileRowProps {
  file: File;
  /** Posição entre as fichas recém-adicionadas (escalona a entrada); `null` = já estava na lista. */
  enterOrder: number | null;
  onRemove: () => void;
}

function FileRow({ file, enterOrder, onRemove }: FileRowProps) {
  // Decidido na montagem: a linha anima uma vez, quando nasce, e não reinicia em re-renders.
  const [order] = useState(enterOrder);
  const entering = order !== null;

  return (
    <li
      data-entering={entering ? "" : undefined}
      style={order !== null ? ({ "--d": `${order * STAGGER_MS}ms` } as CSSProperties) : undefined}
      className={cn("relative flex items-center gap-3 overflow-hidden bg-surface px-3 py-2", entering && "pd-file-in")}
    >
      {entering && <span aria-hidden="true" className="pd-file-sweep pointer-events-none absolute inset-0" />}
      <span className={cn("inline-flex", entering && "pd-file-chip")}>
        <IconChip icon={<IconFile className="h-4 w-4" />} size="sm" tone="success" />
      </span>
      <span className="min-w-0 flex-1 truncate text-sm text-text" title={file.name}>
        {file.name}
      </span>
      <span className="shrink-0 text-xs tabular-nums text-text-subtle">{humanFileSize(file.size)}</span>
      <button
        type="button"
        aria-label={`Remover ${file.name}`}
        onClick={onRemove}
        className="relative flex h-7 w-7 shrink-0 items-center justify-center rounded-control text-text-subtle outline-none transition-colors hover:bg-danger/10 hover:text-danger focus-visible:ring-2 focus-visible:ring-accent/50"
      >
        <IconX className="h-4 w-4" />
      </button>
    </li>
  );
}

/** Arquivos escolhidos, um por linha (nome, tamanho e remover), com resumo
 *  de quantidade/tamanho total — substitui os chips soltos (FileChips).
 *  Cada ficha adicionada entra com um efeito (sobe, o ícone "estoura" e um
 *  feixe de luz a varre); as que já estavam na lista não repetem. */
export function FileList({ files, maxFiles, onRemove, onClearAll, clearAllId }: FileListProps) {
  // Começa com o que já existe ao montar (ex.: voltar à etapa das fichas) para não reanimar tudo.
  const [seen] = useState(() => new WeakSet<File>(files));
  useEffect(() => {
    for (const file of files) seen.add(file);
  }, [files, seen]);

  if (files.length === 0) return null;
  const total = files.reduce((sum, file) => sum + file.size, 0);
  let recentes = 0;

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
          <FileRow
            key={idOf(file)}
            file={file}
            enterOrder={seen.has(file) ? null : recentes++}
            onRemove={() => onRemove(index)}
          />
        ))}
      </ul>
    </div>
  );
}
