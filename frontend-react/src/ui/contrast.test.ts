/// <reference types="node" />
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

// Lido do disco: `css: false` no vitest.config.ts esvazia até o import `?raw` do CSS.
const css = readFileSync(resolve(__dirname, "../index.css"), "utf-8");

function block(selector: RegExp): Record<string, string> {
  const match = css.match(selector);
  if (!match) throw new Error(`Bloco ${selector} não encontrado em index.css`);
  return Object.fromEntries(
    [...match[1].matchAll(/--pd-([a-z0-9-]+):\s*(#[0-9a-fA-F]{6})/g)].map((m) => [m[1], m[2]]),
  );
}

const THEMES = {
  claro: block(/^:root \{([^}]*)\}/m),
  escuro: block(/^\.dark \{([^}]*)\}/m),
};

function channels(hex: string): [number, number, number] {
  return [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16)) as [number, number, number];
}

function luminance(hex: string): number {
  const [r, g, b] = channels(hex).map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function ratio(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}

/** Cor `fg` a `alpha` sobre `bg` (como `bg-accent/10` sobre a superfície). */
function tint(fg: string, bg: string, alpha: number): string {
  const [f, b] = [channels(fg), channels(bg)];
  return "#" + f.map((v, i) => Math.round(v * alpha + b[i] * (1 - alpha)).toString(16).padStart(2, "0")).join("");
}

function token(theme: keyof typeof THEMES, name: string): string {
  const value = THEMES[theme][name];
  if (!value) throw new Error(`Token --pd-${name} não definido no tema ${theme}`);
  return value;
}

const AA = 4.5;

describe.each(Object.keys(THEMES) as (keyof typeof THEMES)[])(
  "contraste WCAG AA (texto pequeno) — tema %s",
  (theme) => {
    const pares: [string, string, string][] = [
      ["text-muted", "surface", "texto de apoio sobre cartões"],
      ["text-subtle", "surface", "detalhes (tamanho de arquivo, etapas) sobre cartões"],
      ["text-subtle", "bg", "títulos de seção sobre o fundo da página"],
      ["accent-text", "surface", "texto/link de destaque sobre cartões"],
    ];

    it.each(pares)("%s sobre %s: %s", (fg, bg) => {
      expect(ratio(token(theme, fg), token(theme, bg))).toBeGreaterThanOrEqual(AA);
    });

    it("accent-text sobre o tom de seleção (bg-accent/10): opção marcada e item ativo do menu", () => {
      const selecionado = tint(token(theme, "accent"), token(theme, "surface"), 0.1);
      expect(ratio(token(theme, "accent-text"), selecionado)).toBeGreaterThanOrEqual(AA);
    });
  },
);
