interface SkipLinkProps {
  targetId?: string;
}

/**
 * "Pular para o conteúdo": a 1ª parada do Tab, visível só com foco. É um
 * <button> e não um <a href="#..."> porque o roteador do app é o hash da URL
 * (`#/cadastro`) — um âncora aqui trocaria de tela em vez de só mover o foco.
 */
export function SkipLink({ targetId = "conteudo" }: SkipLinkProps) {
  return (
    <button
      type="button"
      onClick={() => document.getElementById(targetId)?.focus()}
      className="sr-only focus:not-sr-only focus:fixed focus:left-4 focus:top-3 focus:z-[1200] focus:rounded-control focus:bg-accent focus:px-4 focus:py-2 focus:text-sm focus:font-medium focus:text-accent-fg focus:shadow-pop focus:outline-none"
    >
      Pular para o conteúdo
    </button>
  );
}
