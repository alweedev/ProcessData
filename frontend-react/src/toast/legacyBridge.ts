import { pushToast, type ToastType } from "./toastStore";

/**
 * Abas ainda não migradas (app.v2.js, inativacao/index.js, history/index.js)
 * continuam chamando `showToast(message, type)` normalmente — este módulo só
 * precisa ser importado (efeito colateral no top-level) para a ponte existir.
 * Conteúdo sempre renderiza como texto JSX no ToastViewport (nunca via
 * dangerouslySetInnerHTML), o que corrige o toast.innerHTML sem escaping da
 * implementação legada.
 */
window.showToast = (message: string, type?: string) => {
  // Mesma regra da implementação legada: só "success" (ou ausente) e "info"
  // têm cor própria — qualquer outro valor (danger, warning, error...) virava
  // vermelho lá, então mantemos o mesmo mapeamento aqui.
  let normalized: ToastType;
  if (type === undefined || type === "success") normalized = "success";
  else if (type === "info") normalized = "info";
  else normalized = "danger";
  pushToast(String(message ?? ""), normalized);
};
