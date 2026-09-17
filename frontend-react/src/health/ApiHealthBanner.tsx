import { useState } from "react";
import { Button } from "../ui/Button";
import { TONE_OUTLINE } from "../ui/tone";
import { checkApiHealth, useApiHealth, type ApiState } from "./apiHealthStore";

/**
 * Sempre montado (não condicionalmente): `apiHealthStore` só mantém o
 * polling de 45s rodando enquanto houver ao menos um assinante de
 * `useApiHealth()`. Se este componente só existisse quando `state ===
 * "offline"`, ninguém chamaria o hook enquanto a API estivesse saudável e o
 * app nunca descobriria uma queda.
 */
export function ApiHealthBanner() {
  const state = useApiHealth();

  // Lembra se já esteve offline nesta sessão, pra evitar flicker do banner
  // nos "checking" rotineiros do polling enquanto tudo está saudável — só
  // reaparece "Verificando..." se já tiver caído antes. Ajustado durante o
  // render (padrão React de "adjusting state when a prop changes"), não em
  // efeito — evita uma volta extra de render/commit.
  const [prevState, setPrevState] = useState<ApiState>(state);
  const [everOffline, setEverOffline] = useState(false);
  if (state !== prevState) {
    setPrevState(state);
    if (state === "offline") setEverOffline(true);
    if (state === "online") setEverOffline(false);
  }

  const show = state === "offline" || (state === "checking" && everOffline);
  if (!show) return null;

  return (
    <div role="alert" aria-live="assertive" className={`border-b px-4 py-2 text-sm ${TONE_OUTLINE.danger}`}>
      <div className="mx-auto flex max-w-[1360px] items-center justify-between gap-3">
        <span>
          {state === "checking" ? "Verificando conexão com o servidor…" : "Não foi possível conectar ao servidor."}
        </span>
        <Button variant="outline" size="sm" onClick={() => void checkApiHealth(true)}>
          Tentar novamente
        </Button>
      </div>
    </div>
  );
}
