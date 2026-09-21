/** Uma geração que falhou: a mensagem do servidor e, quando ele separa por arquivo, o motivo de cada um. */
export interface GenerationFailure {
  message: string;
  fileErrors?: Record<string, string>;
}

export interface FailureCause {
  /** arquivo a que o motivo se refere (ausente quando o problema é da geração toda) */
  file?: string;
  text: string;
  /** o que o servidor leu/achou, para conferir sozinho (ex.: as colunas que vieram na planilha) */
  detail?: string;
}

export interface FailureView {
  title: string;
  causes: FailureCause[];
  tips: string[];
  /** a mensagem original, para o suporte (só quando a tela a reescreveu em palavras mais claras) */
  technical?: string;
}

interface Explained {
  text: string;
  detail?: string;
  tips: string[];
  rewritten: boolean;
}

const NO_COLUMNS =
  /^Nenhuma coluna da ficha foi reconhecida \(colunas lidas: (.+?)\)\. O cabeçalho precisa estar na 1ª linha\.(?: Só a 1ª aba \('(.+)'\) é lida\.)?$/;

/** Um motivo do servidor em palavras de quem vai corrigir a ficha, com o que fazer. Desconhecido: mostra como veio. */
function explainCause(text: string): Explained {
  const noColumns = NO_COLUMNS.exec(text);
  if (noColumns) {
    const tips = [
      "Confira se é mesmo a ficha de cadastro (e não outro arquivo, como um mapa de carga já gerado).",
      "O cabeçalho precisa estar na 1ª linha da planilha.",
    ];
    if (noColumns[2]) tips.push(`Só a 1ª aba ("${noColumns[2]}") é lida: deixe os dados nela.`);
    return {
      text: "Não parece uma ficha de cadastro: nenhuma coluna conhecida foi encontrada.",
      detail: `Colunas lidas: ${noColumns[1]}`,
      tips,
      rewritten: true,
    };
  }
  if (text.startsWith("A planilha está vazia")) {
    return {
      text,
      tips: ["Confira se enviou o arquivo certo e se ele tem dados."],
      rewritten: false,
    };
  }
  if (text.startsWith("A planilha não tem linhas de dados")) {
    return {
      text,
      tips: ["Preencha ao menos uma linha abaixo do cabeçalho."],
      rewritten: false,
    };
  }
  if (/nomes? com mais de \d+ caracteres/.test(text)) {
    return {
      text,
      tips: ["Ajuste esses nomes na conferência de nomes e gere de novo."],
      rewritten: false,
    };
  }
  if (text === "Erro de rede.") {
    return {
      text: "Não conseguimos falar com o servidor.",
      tips: ["Verifique sua conexão e tente de novo."],
      rewritten: true,
    };
  }
  const http = /^Erro (\d{3})$/.exec(text);
  if (http) {
    return {
      text: `O servidor teve um problema (erro ${http[1]}).`,
      tips: ["Tente de novo. Se repetir, avise o suporte."],
      rewritten: true,
    };
  }
  return { text, tips: [], rewritten: false };
}

const PREFIXES = ["Nenhum registro processado: ", "Falha ao processar arquivo(s): "];

function withoutPrefix(message: string): string {
  const prefix = PREFIXES.find((p) => message.startsWith(p));
  return prefix ? message.slice(prefix.length) : message;
}

/** Transforma a falha crua ("Nenhum registro processado: a.xls: Nenhuma coluna…") em título, um motivo por arquivo e
 *  as dicas do que fazer, sem repetir a mesma dica por arquivo. */
export function explainFailure(failure: GenerationFailure, title = "Não foi possível gerar o cadastro"): FailureView {
  const perFile = Object.entries(failure.fileErrors ?? {});
  const raw: { file?: string; text: string }[] =
    perFile.length > 0 ? perFile.map(([file, text]) => ({ file, text })) : [{ text: withoutPrefix(failure.message) }];

  const causes: FailureCause[] = [];
  const tips: string[] = [];
  let rewritten = false;
  for (const { file, text } of raw) {
    const explained = explainCause(text);
    causes.push({ file, text: explained.text, detail: explained.detail });
    for (const tip of explained.tips) if (!tips.includes(tip)) tips.push(tip);
    rewritten ||= explained.rewritten;
  }

  return {
    title,
    causes,
    tips,
    technical: rewritten ? failure.message : undefined,
  };
}
