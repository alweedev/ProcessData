import { cn } from "./cn";

interface SkeletonProps {
  className?: string;
}

/** Placeholder de carregamento — bloco com pulso via CSS puro
 *  (@keyframes pd-shimmer em index.css). Puramente decorativo: quem usa
 *  este componente continua anunciando o texto real do estado de loading
 *  via aria-live/sr-only, já que este `<div>` é aria-hidden. */
export function Skeleton({ className }: SkeletonProps) {
  return (
    <div
      aria-hidden="true"
      className={cn("rounded-control bg-border [animation:pd-shimmer_1.4s_ease-in-out_infinite]", className)}
    />
  );
}
