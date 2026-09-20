import { useLayoutEffect, useRef, useState } from "react";
import { cn } from "./cn";
import { IconCheck } from "./icons";

export type StepState = "done" | "current" | "todo";

export interface StepperItem {
  label: string;
  /** Linha de apoio (ex.: "2 arquivos", "CPF · SELF"); some no mobile. */
  detail?: string;
  state: StepState;
  /** Ao alcance do usuário: vira botão (se houver `onSelect` e não for a etapa atual). */
  selectable?: boolean;
}

interface StepperProps {
  steps: StepperItem[];
  label: string;
  /** Torna clicáveis os pontos `selectable`; sem isto a linha do tempo é só informativa. */
  onSelect?: (index: number) => void;
}

const BUBBLE: Record<StepState, string> = {
  done: "border-success/40 bg-success-soft text-success",
  current: "border-accent bg-accent text-accent-fg",
  todo: "border-border-strong bg-surface text-text-subtle",
};

const STATE_HINT: Record<StepState, string> = {
  done: " (concluída)",
  current: " (etapa atual)",
  todo: "",
};

const RING = 40; // diâmetro do anel que acompanha a etapa atual (o ponto tem 32)

/**
 * Linha do tempo do fluxo. Um anel destaca a etapa atual e PULA em arco de um
 * ponto ao outro quando ela muda; a linha entre as etapas se preenche.
 * Com `prefers-reduced-motion` o index.css zera as durações e o anel só troca de lugar.
 */
export function Stepper({ steps, label, onSelect }: StepperProps) {
  const listRef = useRef<HTMLOListElement>(null);
  const dotRefs = useRef<(HTMLSpanElement | null)[]>([]);
  const currentIndex = steps.findIndex((step) => step.state === "current");

  // Centro (x) do ponto atual, relativo à lista. O anel é posicionado por ele.
  const [x, setX] = useState<number | null>(null);
  // Só anima depois da 1ª medição, senão o anel "voaria" da esquerda ao abrir a tela.
  const [animate, setAnimate] = useState(false);

  useLayoutEffect(() => {
    const measure = () => {
      const dot = dotRefs.current[currentIndex];
      const list = listRef.current;
      if (!dot || !list) return setX(null);
      const d = dot.getBoundingClientRect();
      const l = list.getBoundingClientRect();
      setX(d.left - l.left + d.width / 2);
    };
    measure(); // a cada render: o texto dos detalhes muda e desloca os pontos
    if (typeof ResizeObserver === "undefined" || !listRef.current) return;
    const observer = new ResizeObserver(measure);
    observer.observe(listRef.current);
    return () => observer.disconnect();
  });

  useLayoutEffect(() => {
    const frame = requestAnimationFrame(() => setAnimate(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  // Cada mudança de etapa gera um novo id: remonta o anel e reinicia o "pulo".
  // (Ajuste de estado durante o render, padrão do React — evita um efeito extra.)
  const [hop, setHop] = useState({ index: currentIndex, id: 0 });
  if (hop.index !== currentIndex) setHop({ index: currentIndex, id: hop.id + 1 });

  return (
    <ol ref={listRef} aria-label={label} className="relative mb-5 flex items-start">
      {x !== null && (
        <span
          aria-hidden="true"
          data-testid="stepper-marker"
          className="pointer-events-none absolute left-0 z-10"
          style={{
            top: -(RING - 32) / 2,
            width: RING,
            height: RING,
            transform: `translateX(${x - RING / 2}px)`,
            transition: animate ? "transform 550ms cubic-bezier(0.34, 1.25, 0.64, 1)" : "none",
          }}
        >
          <span
            key={hop.id}
            className={cn(
              "block h-full w-full rounded-pill border-2 border-accent-text/70",
              hop.id > 0 && "animate-[pd-hop_550ms_ease-in-out]",
            )}
          />
        </span>
      )}

      {steps.map((step, index) => {
        const interactive = Boolean(onSelect) && Boolean(step.selectable) && step.state !== "current";
        const Wrapper = interactive ? "button" : "div";
        return (
          <li
            key={step.label}
            aria-current={step.state === "current" ? "step" : undefined}
            className={cn("flex items-start", index < steps.length - 1 ? "flex-1" : "flex-none")}
          >
            <Wrapper
              {...(interactive ? { type: "button" as const, onClick: () => onSelect?.(index) } : {})}
              className={cn(
                "group flex min-w-0 items-center gap-2 text-left sm:gap-3",
                interactive &&
                  "cursor-pointer rounded-control outline-none focus-visible:ring-2 focus-visible:ring-accent/50",
              )}
            >
              <span
                ref={(el) => {
                  dotRefs.current[index] = el;
                }}
                aria-hidden="true"
                data-testid={`stepper-dot-${index}`}
                className={cn(
                  "flex h-8 w-8 shrink-0 items-center justify-center rounded-pill border text-sm font-semibold transition-colors",
                  BUBBLE[step.state],
                )}
              >
                {step.state === "done" ? (
                  <IconCheck className="h-4 w-4 animate-[pd-pop_320ms_ease-out]" />
                ) : (
                  index + 1
                )}
              </span>
              {/* Mobile: só a etapa atual mostra o rótulo; as demais ficam só no círculo. */}
              <span className={cn("min-w-0", step.state !== "current" && "sr-only sm:not-sr-only")}>
                <span
                  className={cn(
                    "block truncate text-sm font-medium",
                    step.state === "todo" ? "text-text-muted" : "text-text",
                    interactive && "group-hover:underline",
                  )}
                >
                  {step.label}
                  <span className="sr-only">{STATE_HINT[step.state]}</span>
                </span>
                {step.detail && <span className="hidden truncate text-xs text-text-subtle sm:block">{step.detail}</span>}
              </span>
            </Wrapper>
            {index < steps.length - 1 && (
              <span aria-hidden="true" className="relative mx-2 mt-4 h-px min-w-4 flex-1 bg-border sm:mx-3 sm:min-w-6">
                {/* A linha se preenche quando a etapa seguinte é alcançada. */}
                <span
                  className={cn(
                    "absolute inset-0 origin-left bg-success/60 transition-transform duration-500",
                    steps[index + 1].state === "todo" ? "scale-x-0" : "scale-x-100",
                  )}
                />
              </span>
            )}
          </li>
        );
      })}
    </ol>
  );
}
