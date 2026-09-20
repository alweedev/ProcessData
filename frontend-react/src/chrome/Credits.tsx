import { Badge } from "../ui/Badge";

export function Credits({ className = "" }: { className?: string }) {
  return (
    <div className={`flex flex-wrap items-center gap-x-2 gap-y-1 text-xs text-text-subtle ${className}`}>
      <Badge tone="neutral">v{__APP_VERSION__}</Badge>
      <span>© 2026 ProcessData</span>
      <span className="basis-full">
        Desenvolvido por{" "}
        <a
          href="https://www.linkedin.com/in/alejandro-gabriel/"
          target="_blank"
          rel="noopener noreferrer"
          className="text-accent hover:underline"
        >
          Alejandro Gabriel
        </a>
      </span>
    </div>
  );
}
