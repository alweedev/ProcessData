import { triggerAnchorDownload } from "../lib/downloadFile";
import { useRuns, type RunOperation } from "../runs/runsStore";
import { EmptyState } from "./EmptyState";
import { IconInbox } from "./icons";
import { ResultCard } from "./ResultCard";

const LABEL: Record<RunOperation, string> = {
  cadastro: "Cadastro",
  inativacao: "Inativação",
  estruturas: "Estruturas de aprovação",
};

/** Lista de execuções da sessão para uma operação (com "baixar novamente"). */
export function RunHistoryPanel({ operation, className }: { operation: RunOperation; className?: string }) {
  const runs = useRuns().filter((run) => run.operation === operation);

  return (
    <div className={className ?? "mt-6"}>
      <h3 className="mb-2 text-sm font-semibold uppercase tracking-wide text-text-subtle">Nesta sessão</h3>
      {runs.length === 0 ? (
        <EmptyState
          icon={<IconInbox className="h-5 w-5" />}
          title="Nenhuma execução ainda"
          description="O resultado de cada geração aparece aqui, com opção de baixar de novo."
          compact
        />
      ) : (
        <div className="space-y-2">
          {runs.map((run) => (
            <ResultCard
              key={run.id}
              operationLabel={LABEL[run.operation]}
              timestamp={run.ts}
              inputSummary={run.inputSummary}
              outputFilename={run.outputFilename}
              onRedownload={
                run.blobUrl ? () => triggerAnchorDownload(run.blobUrl as string, run.outputFilename) : undefined
              }
            />
          ))}
        </div>
      )}
    </div>
  );
}
