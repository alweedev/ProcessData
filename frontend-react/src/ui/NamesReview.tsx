import type { NameReviewItem } from "../lib/api";
import {
  MAX_NAME_FIELD,
  applyEdit,
  decisionFor,
  fits,
  isResolved,
  pendingItems,
  sanitizeName,
  type Decisions,
  type NameDecision,
} from "../lib/nameReview";
import { Badge } from "./Badge";
import { Button } from "./Button";
import { Field } from "./Field";
import { IconCheck } from "./icons";
import { TextInput } from "./TextInput";

interface NamesReviewProps {
  items: NameReviewItem[];
  decisions: Decisions;
  onDecision: (key: string, decision: NameDecision) => void;
  onAcceptAll: () => void;
}

function plural(n: number, singular: string, pluralForma: string): string {
  return n === 1 ? singular : pluralForma;
}

/**
 * Nomes que o sistema não teve certeza de separar (ou que passaram de 20 caracteres), para conferir ANTES de gerar.
 * A plataforma exige o nome como no documento: Nome com todos os nomes próprios, Sobrenome com todos os sobrenomes.
 * O que o usuário aceita ou corrige aqui também ensina o vocabulário do sistema, e o Gerar só libera sem pendências.
 */
export function NamesReview({ items, decisions, onDecision, onAcceptAll }: NamesReviewProps) {
  const pending = pendingItems(items, decisions).length;
  const canAccept = items.some((item) => {
    const decision = decisionFor(item, decisions);
    return !decision.confirmed && fits(decision.nome, decision.sobrenome);
  });

  return (
    <section
      id="cadastro_names_review"
      aria-labelledby="cadastro_names_review_title"
      className="mb-4 rounded-surface border border-border bg-surface-2 p-4"
    >
      <header className="mb-3 flex flex-wrap items-start justify-between gap-3">
        <div className="min-w-0">
          <h4 id="cadastro_names_review_title" className="text-sm font-semibold text-text">
            Nomes para conferir ({items.length})
          </h4>
          <p role="status" className="mt-0.5 text-xs text-text-muted">
            {pending === 0
              ? "Todos os nomes foram conferidos. Pode gerar."
              : `${pending} ${plural(pending, "nome ainda precisa", "nomes ainda precisam")} da sua confirmação antes de gerar.`}
          </p>
          <p className="mt-1 text-xs text-text-subtle">
            Nome leva todos os nomes próprios e Sobrenome, todos os sobrenomes, como no documento — até {MAX_NAME_FIELD}{" "}
            caracteres em cada.
          </p>
        </div>
        <Button id="cadastro_names_accept_all" variant="secondary" size="sm" disabled={!canAccept} onClick={onAcceptAll}>
          Aceitar todas as sugestões
        </Button>
      </header>

      <ul aria-label="Nomes para conferir" className="space-y-3">
        {items.map((item) => (
          <NameRow
            key={item.key}
            item={item}
            decision={decisionFor(item, decisions)}
            onDecision={(decision) => onDecision(item.key, decision)}
          />
        ))}
      </ul>
    </section>
  );
}

function NameRow({
  item,
  decision,
  onDecision,
}: {
  item: NameReviewItem;
  decision: NameDecision;
  onDecision: (decision: NameDecision) => void;
}) {
  const id = `cadastro_name_${item.key.replace(":", "-")}`;
  const nomeLength = sanitizeName(decision.nome).length;
  const sobrenomeLength = sanitizeName(decision.sobrenome).length;
  const resolved = isResolved(decision);
  const canConfirm = fits(decision.nome, decision.sobrenome);

  return (
    <li
      data-resolved={resolved}
      className="rounded-control border border-border bg-surface p-3"
      aria-label={`${item.label}: ${item.nome_completo}`}
    >
      <div className="mb-2 flex flex-wrap items-center gap-x-2 gap-y-1">
        <span className="text-xs font-medium text-text-subtle">{item.label}</span>
        <span className="text-xs text-text-muted">no documento:</span>
        <span className="min-w-0 break-words font-mono text-xs text-text">{item.nome_completo}</span>
      </div>

      <div className="mb-2 flex flex-wrap gap-1.5">
        {item.motivos.map((motivo) => (
          <Badge key={motivo} tone="warning">
            {motivo}
          </Badge>
        ))}
        {!canConfirm && (nomeLength > MAX_NAME_FIELD || sobrenomeLength > MAX_NAME_FIELD) && (
          <Badge tone="danger">Passa de {MAX_NAME_FIELD} caracteres: abrevie ou corrija</Badge>
        )}
      </div>

      <div className="grid gap-3 sm:grid-cols-[1fr_1fr_auto] sm:items-start">
        <Field
          id={`${id}_nome`}
          label="Nome"
          hint={`${nomeLength}/${MAX_NAME_FIELD}`}
          error={nomeLength > MAX_NAME_FIELD ? `${nomeLength}/${MAX_NAME_FIELD}: passa do limite` : undefined}
        >
          <TextInput
            id={`${id}_nome`}
            value={decision.nome}
            aria-invalid={nomeLength > MAX_NAME_FIELD}
            onChange={(event) => onDecision(applyEdit(decision, { nome: event.target.value }))}
          />
        </Field>
        <Field
          id={`${id}_sobrenome`}
          label="Sobrenome"
          hint={`${sobrenomeLength}/${MAX_NAME_FIELD}`}
          error={sobrenomeLength > MAX_NAME_FIELD ? `${sobrenomeLength}/${MAX_NAME_FIELD}: passa do limite` : undefined}
        >
          <TextInput
            id={`${id}_sobrenome`}
            value={decision.sobrenome}
            aria-invalid={sobrenomeLength > MAX_NAME_FIELD}
            onChange={(event) => onDecision(applyEdit(decision, { sobrenome: event.target.value }))}
          />
        </Field>
        <div className="flex min-h-9 items-center sm:mt-6">
          {resolved ? (
            <span className="inline-flex items-center gap-1.5 text-sm font-medium text-success">
              <IconCheck className="h-4 w-4" />
              Conferido
            </span>
          ) : (
            <Button
              id={`${id}_confirm`}
              variant="secondary"
              size="sm"
              disabled={!canConfirm}
              onClick={() => onDecision({ ...decision, confirmed: true })}
            >
              Confirmar
            </Button>
          )}
        </div>
      </div>
    </li>
  );
}
