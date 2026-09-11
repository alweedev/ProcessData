import { Kbd } from "../ui/Kbd";
import { Modal } from "./Modal";

const ROWS: { keys: string[]; label: string }[] = [
  { keys: ["H"], label: "Ir para o Início" },
  { keys: ["1"], label: "Abrir Cadastro" },
  { keys: ["2"], label: "Abrir Inativação" },
  { keys: ["3"], label: "Abrir Estruturas de aprovação" },
  { keys: ["4"], label: "Abrir Histórico" },
  { keys: ["Ctrl", "Enter"], label: "Executar a ação principal da tela" },
  { keys: ["?"], label: "Abrir / fechar esta ajuda" },
];

export function ShortcutsHelp({ open, onClose }: { open: boolean; onClose: () => void }) {
  return (
    <Modal open={open} onClose={onClose} title="Atalhos de teclado">
      <ul className="space-y-2">
        {ROWS.map((row) => (
          <li key={row.label} className="flex items-center justify-between gap-4">
            <span className="text-text">{row.label}</span>
            <span className="flex shrink-0 items-center gap-1">
              {row.keys.map((k) => (
                <Kbd key={k}>{k}</Kbd>
              ))}
            </span>
          </li>
        ))}
      </ul>
      <p className="mt-3 text-xs text-text-subtle">
        Os atalhos de navegação ficam inativos enquanto você digita num campo.
      </p>
    </Modal>
  );
}
