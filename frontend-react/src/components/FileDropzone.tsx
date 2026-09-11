import { useEffect, useRef, useState, type ReactNode } from "react";

interface FileDropzoneProps {
  id: string;
  containerId?: string;
  accept: string;
  multiple?: boolean;
  ariaLabel: string;
  description: string;
  onFiles: (files: FileList) => void;
  /** Quando fornecido, o <input> nativo é reconstruído (via DataTransfer) pra
   *  espelhar essa lista — permite "remover 1 arquivo" mantendo `input.files`
   *  em sincronia com o estado do hook. */
  syncFiles?: File[];
  /** Renderizado abaixo do botão "Selecionar" — cada aba cuida do próprio feedback. */
  children?: ReactNode;
}

/**
 * Zona de upload compartilhada (clique, arrastar-e-soltar, ativação por
 * teclado) — unifica o padrão que Cadastro/Inativação/Estruturas
 * reimplementavam de forma independente no app legado.
 */
export function FileDropzone({
  id,
  containerId,
  accept,
  multiple,
  ariaLabel,
  description,
  onFiles,
  syncFiles,
  children,
}: FileDropzoneProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  useEffect(() => {
    if (!syncFiles || !inputRef.current) return;
    const dt = new DataTransfer();
    for (const file of syncFiles) dt.items.add(file);
    inputRef.current.files = dt.files;
  }, [syncFiles]);

  return (
    <div
      id={containerId}
      aria-live="polite"
      tabIndex={0}
      role="button"
      aria-label={ariaLabel}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          inputRef.current?.click();
        }
      }}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        if (e.dataTransfer.files?.length) onFiles(e.dataTransfer.files);
      }}
      className={`rounded-lg border-2 border-dashed p-6 text-center transition-colors ${
        dragOver ? "border-accent bg-accent/5" : "border-border-strong hover:border-accent/50"
      }`}
    >
      <p className="mb-3 text-sm text-text-muted">{description}</p>
      <input
        ref={inputRef}
        type="file"
        id={id}
        accept={accept}
        multiple={multiple}
        aria-label={ariaLabel}
        className="hidden"
        onChange={(e) => {
          if (e.target.files?.length) onFiles(e.target.files);
        }}
      />
      <button
        type="button"
        onClick={() => inputRef.current?.click()}
        className="rounded-lg bg-accent px-4 py-2 text-sm font-medium text-accent-fg transition-colors hover:bg-accent-hover"
      >
        Selecionar
      </button>
      {children}
    </div>
  );
}
