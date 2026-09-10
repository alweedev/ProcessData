import { useRef, useState, type ReactNode } from "react";

interface FileDropzoneProps {
  id: string;
  containerId?: string;
  accept: string;
  multiple?: boolean;
  ariaLabel: string;
  description: string;
  onFiles: (files: FileList) => void;
  /** Renderizado abaixo do botão "Selecionar" — cada aba cuida do próprio feedback (nome do arquivo, chip de status etc.). */
  children?: ReactNode;
}

/**
 * Zona de upload compartilhada (clique, arrastar-e-soltar, ativação por
 * teclado) — unifica o padrão que Cadastro/Inativação/Estruturas
 * reimplementavam de forma independente no app legado.
 */
export function FileDropzone({ id, containerId, accept, multiple, ariaLabel, description, onFiles, children }: FileDropzoneProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

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
      className={`tw:rounded-lg tw:border-2 tw:border-dashed tw:p-6 tw:text-center tw:transition-colors ${
        dragOver ? "tw:border-accent tw:bg-accent/5" : "tw:border-accent/40"
      } tw:dark:border-accent-dark/40`}
    >
      <p className="tw:mb-3 tw:text-sm tw:text-slate-500 tw:dark:text-slate-400">{description}</p>
      <input
        ref={inputRef}
        type="file"
        id={id}
        accept={accept}
        multiple={multiple}
        aria-label={ariaLabel}
        className="tw:hidden"
        onChange={(e) => {
          if (e.target.files?.length) onFiles(e.target.files);
        }}
      />
      <button
        type="button"
        onClick={() => inputRef.current?.click()}
        className="tw:rounded-lg tw:bg-accent tw:px-4 tw:py-2 tw:text-sm tw:font-medium tw:text-white tw:hover:bg-accent-hover"
      >
        Selecionar
      </button>
      {children}
    </div>
  );
}
