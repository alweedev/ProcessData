/**
 * O que ainda falta para liberar "Validar" e "Gerar" no Cadastro, em uma
 * frase para a dica ao lado dos botões — ou `null` quando nada falta.
 * Tipo de login e fluxo não têm valor padrão: precisam ser escolhidos.
 */
export function descreverPendencias(temArquivos: boolean, login: string | null, fluxo: string | null): string | null {
  const faltas: string[] = [];
  if (!temArquivos) faltas.push("enviar ao menos uma ficha");
  if (login === null && fluxo === null) faltas.push("escolher o tipo de login e o fluxo");
  else if (login === null) faltas.push("escolher o tipo de login");
  else if (fluxo === null) faltas.push("escolher o fluxo");

  return faltas.length ? `Para continuar: ${faltas.join(" e ")}.` : null;
}
