import { createContext, useContext, useEffect } from "react";

type ActionFn = () => void;

export interface ActionRegistry {
  register: (fn: ActionFn | null) => void;
  run: () => void;
}

/** Registro da "ação primária da tela ativa" — a casca guarda um ref, cada
 *  tela registra o próprio submit num efeito, e o atalho Ctrl+Enter chama
 *  `run()`. Ref em vez de state pra não re-renderizar a árvore inteira. */
export const ActionContext = createContext<ActionRegistry | null>(null);

export function useActionRunner(): () => void {
  const ctx = useContext(ActionContext);
  return () => ctx?.run();
}

/** Registra `fn` como a ação primária enquanto a tela estiver montada.
 *  Passe `null` quando a ação estiver indisponível (ex.: já rodando). */
export function useRegisterPrimaryAction(fn: ActionFn | null): void {
  const ctx = useContext(ActionContext);
  useEffect(() => {
    if (!ctx) return;
    ctx.register(fn);
    return () => ctx.register(null);
  }, [ctx, fn]);
}
