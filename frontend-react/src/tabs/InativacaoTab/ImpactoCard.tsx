import type { EstruturaComoAprovador, UsuarioAnalise } from "../../lib/inativacaoApi";
import { cn } from "../../ui/cn";

const ROTULO: Record<UsuarioAnalise["situacao"], string> = {
  EXECUTAVEL: "Será inativado",
  SEM_CPF: "Sem CPF no cadastro",
  JA_INATIVO: "Já inativo",
  NAO_LOCALIZADO: "Não localizado",
  PENDENTE_SELECAO: "Escolha o usuário",
};

function descreverPosicao(e: EstruturaComoAprovador): string {
  const partes: string[] = [];
  if (e.posicoes.length === 1) partes.push(`posição ${e.posicoes[0]}`);
  if (e.posicoes.length > 1) partes.push(`posições ${e.posicoes.join(", ")}`);
  if (e.segundoNivel) partes.push("2º nível");
  return partes.join(" · ");
}

interface ImpactoCardProps {
  usuario: UsuarioAnalise;
  /** CPFs marcados entre os candidatos de homônimos. */
  escolhidos: string[];
  onEscolher: (cpf: string, marcado: boolean) => void;
  /** Trava a escolha de homônimos (ex.: enquanto uma análise está em andamento). */
  disabled?: boolean;
}

/** Um usuário da análise: quem é, a situação e o que muda nas estruturas de aprovação. */
export function ImpactoCard({ usuario, escolhidos, onEscolher, disabled = false }: ImpactoCardProps) {
  const executavel = usuario.situacao === "EXECUTAVEL";
  const orfas = usuario.comoAprovador.filter((e) => e.acao === "ORFA").length;
  const identificacao = [usuario.cpfMascarado, usuario.email].filter(Boolean).join(" · ");

  return (
    <li
      data-situacao={usuario.situacao}
      className={cn(
        "rounded-control border px-4 py-3 text-sm",
        executavel ? "border-border bg-surface" : "border-warning/40 bg-warning-soft",
      )}
    >
      <div className="flex flex-wrap items-baseline justify-between gap-2">
        <p className="font-medium text-text">{usuario.nome || "—"}</p>
        <span
          className={cn(
            "rounded-pill border px-2 py-0.5 text-xs font-medium",
            executavel ? "border-success/40 bg-success-soft text-success" : "border-warning/40 text-warning",
          )}
        >
          {ROTULO[usuario.situacao]}
        </span>
      </div>
      {identificacao && <p className="mt-0.5 text-xs text-text-muted">{identificacao}</p>}
      {usuario.alerta && <p className="mt-2 text-warning">{usuario.alerta}</p>}

      {usuario.situacao === "PENDENTE_SELECAO" && (
        <fieldset className="mt-2">
          <legend className="text-text-muted">Há mais de um usuário com este nome. Marque quem deve ser inativado:</legend>
          {usuario.candidatos.map((c) => (
            <label key={`${c.cpf}-${c.email}`} className="mt-1.5 flex items-start gap-2 text-text">
              <input
                type="checkbox"
                className="mt-0.5 h-4 w-4"
                disabled={!c.cpf || disabled}
                checked={escolhidos.includes(c.cpf)}
                onChange={(e) => onEscolher(c.cpf, e.target.checked)}
              />
              <span>
                {c.nome} · {c.cpfMascarado || "sem CPF"} · {c.email || "sem e-mail"}
              </span>
            </label>
          ))}
        </fieldset>
      )}

      {executavel && (
        <dl className="mt-2 space-y-2">
          <div>
            <dt className="text-xs font-medium text-text-muted">Estrutura do viajante</dt>
            <dd className="text-text">
              {usuario.estruturasViajante.length > 0
                ? `Será excluída: ${usuario.estruturasViajante.join(", ")}`
                : "Nenhuma estrutura direta encontrada."}
            </dd>
          </div>
          <div>
            <dt className="text-xs font-medium text-text-muted">Como aprovador</dt>
            <dd className="text-text">
              {usuario.comoAprovador.length === 0 ? (
                "Não aparece como aprovador em outras estruturas."
              ) : (
                <ul className="space-y-1">
                  {usuario.comoAprovador.map((e) => (
                    <li key={e.aprovacaoId} className="flex flex-wrap items-center gap-2">
                      <span className="font-mono">{e.aprovacaoId}</span>
                      <span className="text-text-muted">{descreverPosicao(e)}</span>
                      <span
                        className={cn(
                          "rounded-pill border px-2 py-0.5 text-xs font-medium",
                          e.acao === "ORFA"
                            ? "border-danger/40 bg-danger-soft text-danger"
                            : "border-warning/40 bg-warning-soft text-warning",
                        )}
                      >
                        {e.acao === "ORFA" ? "Estrutura órfã" : "Compactação"}
                      </span>
                    </li>
                  ))}
                </ul>
              )}
              {orfas > 0 && (
                <p className="mt-1.5 text-danger">
                  Atenção: {orfas === 1 ? "1 estrutura ficará" : `${orfas} estruturas ficarão`} sem nenhum aprovador na
                  Argo.
                </p>
              )}
            </dd>
          </div>
        </dl>
      )}
    </li>
  );
}
