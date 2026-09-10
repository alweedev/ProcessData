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
      className="fixed inset-0 z-[1200] flex items-center justify-center bg-black/50 p-4"
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
        className="w-full max-w-lg rounded-xl border border-black/10 bg-surface shadow-xl outline-none dark:border-white/10 dark:bg-surface-dark"
      >
        <div className="flex items-center justify-between border-b border-black/10 px-4 py-3 dark:border-white/10">
          <h3 className="text-base font-semibold text-slate-900 dark:text-slate-100">{title}</h3>
          <button
            type="button"
            aria-label="Fechar"
            onClick={onClose}
            className="text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-100"
          >
            ✕
          </button>
        </div>
        <div className="px-4 py-4 text-sm text-slate-700 dark:text-slate-200">{children}</div>
        <div className="flex justify-end gap-2 border-t border-black/10 px-4 py-3 dark:border-white/10">
          {footer ?? (
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg bg-accent px-4 py-1.5 text-sm font-medium text-white hover:bg-accent-hover"
            >
              Fechar
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
