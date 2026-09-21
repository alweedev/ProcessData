import type { FailureView } from "../lib/failure";
import { IconAlert } from "./icons";
import { TONE_OUTLINE } from "./tone";

interface GenerationErrorProps {
  view: FailureView;
  /** Espaço em volta: quem usa o cartão decide (vem depois de uma lista ou antes dos botões). */
  className?: string;
  /** Um cartão por vez na tela usa o id padrão; a falha da validação usa outro. */
  id?: string;
}

/** A validação ou a geração falhou: o que aconteceu (um motivo por arquivo), o que fazer e, para o suporte, a mensagem original. */
export function GenerationError({ view, id = "cadastro_debug", className }: GenerationErrorProps) {
  return (
    <div
      id={id}
      aria-live="assertive"
      className={`flex items-start gap-3 rounded-control border px-4 py-3 text-sm ${TONE_OUTLINE.danger} ${className ?? ""}`}
    >
      <IconAlert className="mt-0.5 h-5 w-5 shrink-0" />
      <div className="min-w-0 flex-1 text-text">
        <p className="font-medium text-danger">{view.title}</p>

        <ul className="mt-1.5 space-y-1.5">
          {view.causes.map((cause, index) => (
            <li key={`${cause.file ?? ""}-${index}`}>
              {cause.file && <strong className="font-medium">{cause.file}: </strong>}
              {cause.text}
              {cause.detail && <p className="mt-0.5 break-words text-xs text-text-muted">{cause.detail}</p>}
            </li>
          ))}
        </ul>

        {view.tips.length > 0 && (
          <div className="mt-2.5">
            <p className="text-xs font-medium text-text-subtle">O que fazer</p>
            <ul className="mt-0.5 list-disc space-y-0.5 pl-5 text-text-muted">
              {view.tips.map((tip) => (
                <li key={tip}>{tip}</li>
              ))}
            </ul>
          </div>
        )}

        {view.technical && (
          <details className="mt-2.5 text-xs text-text-muted">
            <summary className="cursor-pointer select-none">Detalhes técnicos</summary>
            <p className="mt-1 break-words font-mono">{view.technical}</p>
          </details>
        )}
      </div>
    </div>
  );
}
