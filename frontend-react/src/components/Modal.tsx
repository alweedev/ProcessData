import { useEffect, useRef, type ReactNode } from "react";

interface ModalProps {
  open: boolean;
  title: string;
  onClose: () => void;
  children: ReactNode;
  /** Rodapé customizado (botões). Se ausente, mostra só um "Fechar". */
  footer?: ReactNode;
}

/**
 * Modal simples (overlay + painel, fecha no Esc e no clique fora).
 * Substitui os `Swal.fire(...)` que sobraram — sem dependência externa.
 */
export function Modal({ open, title, onClose, children, footer }: ModalProps) {
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    panelRef.current?.focus();
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") onClose();
    }
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;

  return (
    <div
      className="tw:fixed tw:inset-0 tw:z-[1200] tw:flex tw:items-center tw:justify-center tw:bg-black/50 tw:p-4"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        tabIndex={-1}
        className="tw:w-full tw:max-w-lg tw:rounded-xl tw:border tw:border-black/10 tw:bg-surface tw:shadow-xl tw:outline-none tw:dark:border-white/10 tw:dark:bg-surface-dark"
      >
        <div className="tw:flex tw:items-center tw:justify-between tw:border-b tw:border-black/10 tw:px-4 tw:py-3 tw:dark:border-white/10">
          <h3 className="tw:text-base tw:font-semibold tw:text-slate-900 tw:dark:text-slate-100">{title}</h3>
          <button
            type="button"
            aria-label="Fechar"
            onClick={onClose}
            className="tw:text-slate-500 tw:hover:text-slate-800 tw:dark:text-slate-400 tw:dark:hover:text-slate-100"
          >
            ✕
          </button>
        </div>
        <div className="tw:px-4 tw:py-4 tw:text-sm tw:text-slate-700 tw:dark:text-slate-200">{children}</div>
        <div className="tw:flex tw:justify-end tw:gap-2 tw:border-t tw:border-black/10 tw:px-4 tw:py-3 tw:dark:border-white/10">
          {footer ?? (
            <button
              type="button"
              onClick={onClose}
              className="tw:rounded-lg tw:bg-accent tw:px-4 tw:py-1.5 tw:text-sm tw:font-medium tw:text-white tw:hover:bg-accent-hover"
            >
              Fechar
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
