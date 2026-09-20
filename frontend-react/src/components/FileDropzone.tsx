import { useEffect, useRef, useState, type ReactNode } from "react";
import { Button } from "../ui/Button";
import { cn } from "../ui/cn";
import { IconChip } from "../ui/IconChip";
import { IconUpload } from "../ui/icons";

interface FileDropzoneProps {
  id: string;
  containerId?: string;
  accept: string;
  multiple?: boolean;
  ariaLabel: string;
  description: string;
  /** Linha de apoio sob a descrição (formatos e limites) — só na variante `rich`. */
  hint?: string;
  /** `classic` (padrão): layout original. `rich`: área inteira clicável, ícone,
   *  destaque ao arrastar e texto de apoio. */
  variant?: "classic" | "rich";
  /** Só `rich`: versão enxuta (sem ícone/hint) para quando já há arquivos escolhidos. */
  compact?: boolean;
  /** Texto do botão (padrão "Selecionar"). */
  buttonLabel?: string;
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
  hint,
  variant = "classic",
  compact = false,
  buttonLabel = "Selecionar",
  onFiles,
  syncFiles,
  children,
}: FileDropzoneProps) {
  const rich = variant === "rich";
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
      onDragLeave={(e) => {
        // Ignora o "leave" ao passar de um filho para outro (evita piscar).
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setDragOver(false);
      }}
      onClick={
        rich
          ? (e) => {
              // Cliques em botões/inputs internos (Selecionar, remover...) têm ação própria.
              if ((e.target as HTMLElement).closest("button, a, input, select, textarea")) return;
              inputRef.current?.click();
            }
          : undefined
      }
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        if (e.dataTransfer.files?.length) onFiles(e.dataTransfer.files);
      }}
      className={cn(
        "border-2 border-dashed text-center transition",
        rich ? cn("cursor-pointer rounded-surface px-6", compact ? "py-4" : "py-8") : "rounded-control p-6",
        dragOver
          ? "border-accent bg-accent/10"
          : rich
            ? "border-border-strong bg-surface-2/50 hover:border-accent/60 hover:bg-accent/5"
            : "border-border-strong hover:border-accent/50",
        rich && dragOver && "scale-[1.01]",
      )}
    >
      {rich && !compact && (
        <div className={cn("mx-auto mb-3 w-fit transition-transform", dragOver && "-translate-y-1")}>
          <IconChip icon={<IconUpload className="h-5 w-5" />} size="lg" shape="circle" />
        </div>
      )}
      <p className={rich ? "text-sm font-medium text-text" : "mb-3 text-sm text-text-muted"}>
        {rich && dragOver ? "Solte para enviar" : description}
      </p>
      {rich && hint && !compact && <p className="mb-4 mt-1 text-xs text-text-muted">{hint}</p>}
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
      <Button
        type="button"
        variant={compact ? "secondary" : "primary"}
        className={compact ? "mt-3" : undefined}
        onClick={() => inputRef.current?.click()}
      >
        {buttonLabel}
      </Button>
      {children}
    </div>
  );
}
