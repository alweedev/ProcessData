import type { CSSProperties, ReactNode } from "react";
import { useApiHealth, type ApiState } from "../../health/apiHealthStore";
import { navigate, type View } from "../../routing/useHashRoute";
import { Badge } from "../../ui/Badge";
import { Card } from "../../ui/Card";
import { EmptyState } from "../../ui/EmptyState";
import { IconChip } from "../../ui/IconChip";
import { IconClock, IconInbox, IconSitemap, IconUpload, IconUserMinus } from "../../ui/icons";
import { PageHeader } from "../../ui/PageHeader";
import { ResultCard } from "../../ui/ResultCard";
import { useRecentRuns } from "./useRecentRuns";

interface Op {
  id: Exclude<View, "home">;
  label: string;
  desc: string;
  icon: ReactNode;
}

const OPS: Op[] = [
  {
    id: "cadastro",
    label: "Cadastro em massa",
    desc: "Trate a planilha de fichas e gere o arquivo de carga.",
    icon: <IconUpload className="h-5 w-5" />,
  },
  {
    id: "inativacao",
    label: "Inativação",
    desc: "Busque usuários numa base e gere o arquivo de inativação.",
    icon: <IconUserMinus className="h-5 w-5" />,
  },
  {
    id: "estruturas",
    label: "Estruturas de aprovação",
    desc: "Remova um aprovador e recompacte as estruturas.",
    icon: <IconSitemap className="h-5 w-5" />,
  },
  {
    id: "historico",
    label: "Histórico",
    desc: "Consulte e exporte as execuções recentes.",
    icon: <IconClock className="h-5 w-5" />,
  },
];

const API_TEXT: Record<ApiState, string> = {
  online: "Online",
  offline: "Offline",
  checking: "Verificando…",
};
const API_TONE: Record<ApiState, "success" | "danger" | "neutral"> = {
  online: "success",
  offline: "danger",
  checking: "neutral",
};

export function HomeView() {
  const api = useApiHealth();
  const recent = useRecentRuns(6);

  return (
    <div>
      <PageHeader
        title="Visão geral"
        description="Ferramentas de tratamento de planilhas do suporte Vermari."
        actions={<Badge tone={API_TONE[api]}>API {API_TEXT[api]}</Badge>}
      />

      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
        {OPS.map((op, index) => (
          <Card
            key={op.id}
            interactive
            role="button"
            tabIndex={0}
            className="pd-enter pd-stagger"
            style={{ "--i": index } as CSSProperties}
            onClick={() => navigate(op.id)}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                navigate(op.id);
              }
            }}
          >
            <div className="flex items-start gap-3">
              <IconChip icon={op.icon} />
              <div>
                <p className="font-medium text-text">{op.label}</p>
                <p className="mt-0.5 text-sm text-text-muted">{op.desc}</p>
              </div>
            </div>
          </Card>
        ))}
      </div>

      <h3 className="mb-3 mt-8 text-sm font-semibold uppercase tracking-wide text-text-subtle">Últimas execuções</h3>
      {recent.length === 0 ? (
        <EmptyState
          icon={<IconInbox className="h-5 w-5" />}
          title="Nada por aqui ainda"
          description="As execuções que você rodar aparecem aqui e no Histórico."
          compact
        />
      ) : (
        <div className="space-y-2">
          {recent.map((item, index) => (
            <div key={item.key} className="pd-enter pd-stagger" style={{ "--i": index } as CSSProperties}>
              <ResultCard
                compact
                operationLabel={item.title}
                timestamp={item.ts}
                inputSummary={item.subtitle ?? ""}
                outputFilename={item.outputFilename}
                status={item.status}
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
