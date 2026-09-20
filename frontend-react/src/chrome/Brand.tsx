import { navigate } from "../routing/useHashRoute";

interface BrandProps {
  id?: string;
  className?: string;
}

export function Brand({ id, className = "" }: BrandProps) {
  return (
    <button
      id={id}
      type="button"
      onClick={() => navigate("home")}
      aria-label="ProcessData — ir para o início"
      className={`flex items-center gap-2 rounded-control outline-none focus-visible:ring-2 focus-visible:ring-accent/50 ${className}`}
    >
      <span
        aria-hidden="true"
        className="flex h-8 w-8 items-center justify-center rounded-control bg-gradient-to-br from-brand-from to-brand-to text-white"
      >
        <span className="text-xs font-extrabold leading-none tracking-tighter">PD</span>
      </span>
      <span className="text-base font-bold tracking-tight text-text">ProcessData</span>
    </button>
  );
}
