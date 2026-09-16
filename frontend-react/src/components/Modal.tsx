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
 * Sem dependência externa.
 */
const FOCUSABLE_SELECTOR =
  'a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])';

export function Modal({ open, title, onClose, children, footer }: ModalProps) {
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    panelRef.current?.focus();

    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") {
        onClose();
        return;
      }
      if (e.key !== "Tab") return;

      // Focus trap: Tab/Shift+Tab não pode sair do diálogo enquanto ele
      // estiver aberto (o conteúdo por trás fica visualmente coberto, mas
      // sem isso o foco de teclado continua acessível nele).
      const panel = panelRef.current;
      if (!panel) return;
      const focusable = Array.from(panel.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter(
        (el) => el.offsetParent !== null,
      );
      if (focusable.length === 0) {
        e.preventDefault();
        return;
      }
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      const active = document.activeElement;
      if (e.shiftKey) {
        if (active === first || active === panel) {
          e.preventDefault();
          last.focus();
        }
      } else if (active === last) {
        e.preventDefault();
        first.focus();
      }
    }
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;

  return (
    <div
      className="fixed inset-0 z-[1200] flex items-center justify-center bg-black/50 p-4 backdrop-blur-sm"
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
        className="w-full max-w-lg rounded-xl border border-border bg-surface shadow-pop outline-none"
      >
        <div className="flex items-center justify-between border-b border-border px-4 py-3">
          <h3 className="text-base font-semibold text-text">{title}</h3>
          <button
            type="button"
            aria-label="Fechar"
            onClick={onClose}
            className="rounded p-1 text-text-subtle transition-colors hover:bg-surface-2 hover:text-text"
          >
            ✕
          </button>
        </div>
        <div className="px-4 py-4 text-sm text-text-muted">{children}</div>
        <div className="flex justify-end gap-2 border-t border-border px-4 py-3">
          {footer ?? (
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg bg-accent px-4 py-1.5 text-sm font-medium text-accent-fg hover:bg-accent-hover"
            >
              Fechar
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
