# Redesign dark-first + animações Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tornar o frontend dark-first (visual alto-contraste/minimalista estilo Vercel/GitHub-dark), com dark como tema padrão, e adicionar animações mais perceptíveis (entrada de conteúdo, stagger, transição entre telas, loading) usando só CSS nativo e APIs de browser — sem nova dependência JS.

**Architecture:** Todas as mudanças de cor/sombra passam pelos tokens já existentes em `index.css` (indireção `--color-x: var(--pd-x)`, trocando só os valores de `:root`/`.dark`). Animação usa três mecanismos nativos: `@starting-style` (entrada de conteúdo), `transition-delay` via custom property `--i` (stagger), e a View Transitions API já usada no `ThemeToggle` (agora também na troca de aba). Nenhum componente ganha nova prop pública além do necessário.

**Tech Stack:** React 19, Tailwind CSS v4 (tokens via `@theme`/custom properties), Vite 8, Vitest 5. Sem biblioteca de animação nova.

**Spec:** `docs/superpowers/specs/2026-09-17-dark-first-redesign-design.md`

## Global Constraints

- Sem dependência JS nova (nada de `motion`/Framer Motion/etc.) — só CSS + View Transitions API nativa.
- Dark vira o tema padrão quando não há preferência salva nem `prefers-color-scheme: light` explícito; light continua existindo e acessível via `ThemeToggle`.
- Tudo que anima deve respeitar `prefers-reduced-motion: reduce` (gate global já existe em `frontend-react/src/index.css`; código novo em JS que chama `document.startViewTransition` deve checar isso explicitamente, mesmo padrão do `ThemeToggle` atual).
- Sem gestos (drag/swipe/pan) — fora de escopo.
- Light mode não é removido; o `ThemeToggle` continua funcionando como está.

---

## Task 1: Tema escuro como padrão

**Files:**
- Create: `frontend-react/src/chrome/theme.ts`
- Create: `frontend-react/src/chrome/theme.test.ts`
- Modify: `frontend-react/src/chrome/ThemeToggle.tsx` (remove `THEME_KEY`/`Mode`/`computeInitialTheme` locais, linhas 1-18, e importar de `./theme`)
- Modify: `frontend/index.html:14-23` (script anti-FOUC)

**Interfaces:**
- Produces: `computeInitialTheme(): Mode` e `THEME_KEY: string`, exportados de `frontend-react/src/chrome/theme.ts`, consumidos por `ThemeToggle.tsx`.

- [ ] **Step 1: Escrever o teste (falhando)**

Criar `frontend-react/src/chrome/theme.test.ts`:

```ts
import { afterEach, describe, expect, it, vi } from "vitest";
import { computeInitialTheme, THEME_KEY } from "./theme";

function mockMatchMedia(prefersLight: boolean) {
  window.matchMedia = vi.fn().mockImplementation((query: string) => ({
    matches: query === "(prefers-color-scheme: light)" ? prefersLight : false,
    media: query,
    addEventListener: () => {},
    removeEventListener: () => {},
  })) as unknown as typeof window.matchMedia;
}

describe("computeInitialTheme", () => {
  afterEach(() => {
    localStorage.clear();
  });

  it("usa o valor salvo no localStorage quando existe", () => {
    localStorage.setItem(THEME_KEY, "light");
    mockMatchMedia(false);
    expect(computeInitialTheme()).toBe("light");
  });

  it("sem preferência salva e sem prefers-color-scheme:light, usa dark por padrão", () => {
    mockMatchMedia(false);
    expect(computeInitialTheme()).toBe("dark");
  });

  it("sem preferência salva mas com prefers-color-scheme:light, respeita o SO", () => {
    mockMatchMedia(true);
    expect(computeInitialTheme()).toBe("light");
  });
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `cd frontend-react && npx vitest run src/chrome/theme.test.ts`
Expected: FAIL — `theme.ts` ainda não existe (`Cannot find module './theme'`).

- [ ] **Step 3: Criar `frontend-react/src/chrome/theme.ts`**

```ts
export const THEME_KEY = "pd_theme";

export type Mode = "light" | "dark";

/** Tema padrão quando não há preferência salva: dark, a menos que o SO peça
 *  explicitamente light (prefers-color-scheme: light). Mesma lógica do
 *  script anti-FOUC em frontend/index.html — os dois precisam ficar em
 *  sincronia pra não haver flash de tema errado no 1º paint. */
export function computeInitialTheme(): Mode {
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === "dark" || stored === "light") return stored;
  } catch {
    /* localStorage indisponível (modo privado etc.) */
  }
  const prefersLight = window.matchMedia?.("(prefers-color-scheme: light)").matches;
  return prefersLight ? "light" : "dark";
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `cd frontend-react && npx vitest run src/chrome/theme.test.ts`
Expected: PASS (3 testes).

- [ ] **Step 5: Atualizar `ThemeToggle.tsx` pra usar o módulo novo**

Em `frontend-react/src/chrome/ThemeToggle.tsx`, substituir as linhas 1-18 (imports + `THEME_KEY`/`Mode`/`computeInitialTheme` locais) por:

```ts
import { useEffect, useState } from "react";
import { cn } from "../ui/cn";
import { IconMoon, IconSun } from "../ui/icons";
import { computeInitialTheme, THEME_KEY, type Mode } from "./theme";
```

O resto do arquivo (função `applyTheme`, componente `ThemeToggle`) fica igual — só passa a usar os símbolos importados em vez dos locais.

- [ ] **Step 6: Sincronizar o script anti-FOUC**

Em `frontend/index.html`, dentro do `<script>` das linhas 14-23, trocar:

```js
var prefers = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
var mode = stored ? stored : prefers ? "dark" : "light";
```

por:

```js
var prefersLight = window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches;
var mode = stored ? stored : prefersLight ? "light" : "dark";
```

- [ ] **Step 7: Verificar tudo**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test`
Expected: tudo passa, incluindo os 3 testes novos (30 no total).

- [ ] **Step 8: Checar manualmente (sem localStorage salvo)**

Abrir o app com `localStorage.clear()` no console (ou aba anônima) e recarregar — deve abrir em dark por padrão, sem flash de tela clara antes.

- [ ] **Step 9: Commit**

```bash
git add frontend-react/src/chrome/theme.ts frontend-react/src/chrome/theme.test.ts frontend-react/src/chrome/ThemeToggle.tsx frontend/index.html
git commit -m "feat(frontend): dark vira o tema padrão quando não há preferência salva"
```

---

## Task 2: Navegação com View Transition

**Files:**
- Modify: `frontend-react/src/routing/useHashRoute.ts` (reescreve por completo — arquivo pequeno, ~28 linhas)
- Create: `frontend-react/src/routing/useHashRoute.test.ts`

**Interfaces:**
- Consumes: nenhuma (é a base de roteamento).
- Produces: `navigate(view: View): void` e `useHashRoute(): View` — **assinaturas públicas inalteradas**, `App.tsx` e `views/Home/index.tsx` continuam chamando exatamente como hoje, sem editar esses arquivos.

Hoje `useHashRoute` usa `useState` + listener de `hashchange` dentro do componente. Isso não dá pra envolver numa View Transition porque `window.location.hash = x` dispara `hashchange` de forma assíncrona — quando o `startViewTransition` tira o "screenshot" de depois, o React ainda não recomitou. A correção é virar um external store (mesmo padrão já usado em `frontend-react/src/health/apiHealthStore.ts`) com estado em módulo, pra poder forçar o commit síncrono com `flushSync` dentro do callback da transição.

- [ ] **Step 1: Escrever o teste (falhando)**

Criar `frontend-react/src/routing/useHashRoute.test.ts`:

```ts
import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { navigate, useHashRoute } from "./useHashRoute";

describe("useHashRoute", () => {
  // Este projeto roda vitest sem `test.globals: true` (ver vitest.config.ts),
  // então o auto-cleanup do Testing Library não se registra sozinho — sem
  // isso, os hooks montados por `renderHook` vazam entre os testes deste
  // arquivo (nenhum teste existente até aqui usava render/renderHook).
  afterEach(() => {
    cleanup();
  });

  it("navigate() atualiza o hash e notifica o hook (sem View Transitions API)", () => {
    const { result } = renderHook(() => useHashRoute());

    act(() => {
      navigate("inativacao");
    });
    expect(window.location.hash).toBe("#/inativacao");
    expect(result.current).toBe("inativacao");

    act(() => {
      navigate("home");
    });
    expect(window.location.hash).toBe("#/");
    expect(result.current).toBe("home");
  });

  it("navegar duas vezes pra mesma view não quebra nem duplica notificação", () => {
    const { result } = renderHook(() => useHashRoute());
    act(() => {
      navigate("estruturas");
    });
    act(() => {
      navigate("estruturas");
    });
    expect(result.current).toBe("estruturas");
    expect(window.location.hash).toBe("#/estruturas");
  });
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `cd frontend-react && npx vitest run src/routing/useHashRoute.test.ts`
Expected: FAIL (jsdom não dispara `hashchange` de forma síncrona com o `useState` atual, então `result.current` não atualiza dentro do `act()` sem um re-render disparado por evento — o teste não passa com a implementação de hoje).

- [ ] **Step 3: Reescrever `useHashRoute.ts`**

Substituir o conteúdo inteiro de `frontend-react/src/routing/useHashRoute.ts` por:

```ts
import { useSyncExternalStore } from "react";
import { flushSync } from "react-dom";

export type View = "home" | "cadastro" | "inativacao" | "estruturas" | "historico";

const VIEWS: readonly View[] = ["home", "cadastro", "inativacao", "estruturas", "historico"];

function parseHash(): View {
  const raw = window.location.hash.replace(/^#\/?/, "").trim().toLowerCase();
  return (VIEWS as readonly string[]).includes(raw) ? (raw as View) : "home";
}

type Listener = () => void;
const listeners = new Set<Listener>();
let view: View = parseHash();

function setView(next: View) {
  if (next === view) return;
  view = next;
  for (const listener of listeners) listener();
}

window.addEventListener("hashchange", () => setView(parseHash()));

/** Navega trocando o hash (`#/cadastro`). O router é só o hash — nunca toca
 *  no path, então o Flask nunca vê essas URLs e os testes e2e (que fazem
 *  `goto("/")` + clique no trigger) não são afetados.
 *
 *  Quando disponível (e sem prefers-reduced-motion), envolve a troca numa
 *  View Transition nativa — mesmo padrão do ThemeToggle — pra uma transição
 *  suave entre telas. `flushSync` garante que o React já comitou a nova
 *  view antes do browser tirar o "screenshot" de depois; sem isso a
 *  transição captura o estado antigo nos dois lados. A classe `vt-nav` em
 *  `<html>` deixa o CSS (index.css) diferenciar essa transição da do
 *  ThemeToggle (crossfade simples), removida quando a transição termina. */
export function navigate(next: View): void {
  const target = next === "home" ? "#/" : `#/${next}`;
  const commit = () => {
    if (window.location.hash !== target) window.location.hash = target;
    setView(next);
  };

  const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
  if (!reduced && typeof document.startViewTransition === "function") {
    document.documentElement.classList.add("vt-nav");
    const transition = document.startViewTransition(() => flushSync(commit));
    transition.finished.finally(() => document.documentElement.classList.remove("vt-nav"));
  } else {
    commit();
  }
}

function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function getSnapshot(): View {
  return view;
}

export function useHashRoute(): View {
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `cd frontend-react && npx vitest run src/routing/useHashRoute.test.ts`
Expected: PASS (2 testes). Em jsdom, `document.startViewTransition` não existe, então `navigate()` cai no branch `else { commit(); }` — é exatamente esse caminho que o teste exercita.

- [ ] **Step 5: Rodar a suíte inteira**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test`
Expected: tudo passa (32 testes no total). Se `tsc` reclamar de `document.startViewTransition`/`ViewTransition` não existir no tipo de `Document`, é porque `lib` do `tsconfig.app.json` (`["ES2023", "DOM"]`) ainda não cobre esse tipo — **isso já era usado sem erro em `ThemeToggle.tsx` antes desta mudança**, então se compilava lá, compila aqui também (mesma API, mesmo arquivo de config).

- [ ] **Step 6: Checar manualmente**

Rodar `npm run dev`, abrir o app, e trocar de aba pela sidebar (Início → Cadastro → Inativação etc.) — deve ter uma transição sutil (fade + leve deslocamento vertical) em vez de troca instantânea. Testar também com "Emulate CSS prefers-reduced-motion: reduce" no DevTools — a troca deve voltar a ser instantânea.

- [ ] **Step 7: Commit**

```bash
git add frontend-react/src/routing/useHashRoute.ts frontend-react/src/routing/useHashRoute.test.ts
git commit -m "feat(frontend): troca de aba usa View Transitions API nativa"
```

---

## Task 3: Paleta dark-first + tom outline em vez de fill

**Files:**
- Modify: `frontend-react/src/index.css:96-127` (bloco `.dark`) e `frontend-react/src/index.css:64-92` (bloco `:root`)
- Modify: `frontend-react/src/ui/Badge.tsx` (arquivo inteiro, ~20 linhas)
- Modify: `frontend-react/src/ui/ValidationReport.tsx` (2 trechos: bloco de `general_errors` e pills de `blanks`)

Sem teste automatizado — mudança puramente visual. Verificação por `tsc`/`lint`/`test` (nada de lógica muda) + inspeção manual no browser.

**Por que `Button.tsx`/`Card.tsx`/`Modal.tsx`/`ToastViewport.tsx`/`ThemeToggle.tsx` não aparecem em nenhuma task deste plano:** todos já consomem os tokens semânticos (`bg-surface`, `border-border-strong`, `shadow-card`, etc.) em vez de cor/sombra hardcoded — herdam a paleta nova (Task 3) e a sombra nova (Task 4) automaticamente, sem precisar editar o arquivo. `IconChip.tsx` fica de fora de propósito: ele é um selo decorativo de ícone (preenchido), não um indicador de status — continua em `TONE_SOFT` (fill), só `Badge`/`ValidationReport` (que são indicadores de status/tom) migram pra `TONE_OUTLINE`.

- [ ] **Step 1: Ajustar tokens de superfície/borda no `.dark`**

Em `frontend-react/src/index.css`, dentro do bloco `.dark` (linhas 96-127), trocar as 6 linhas de superfície/borda:

```css
  --pd-bg: #0a0d12;
  --pd-surface: #111820;
  --pd-surface-2: #19232f;
  --pd-surface-sunken: #0c1219;
  --pd-border: #202b38;
  --pd-border-strong: #2d3d4e;
```

por:

```css
  --pd-bg: #0a0d12;
  --pd-surface: #0e141b;
  --pd-surface-2: #141b24;
  --pd-surface-sunken: #0a0d12;
  --pd-border: #1f2830;
  --pd-border-strong: #384656;
```

(Menos salto de matiz entre `bg`/`surface`/`surface-2` — `surface-sunken` passa a ser igual a `bg`, já que quem define profundidade agora é a borda, não mais um fundo mais escuro. `border-strong` fica bem mais forte, pra Card/Modal se lerem por contorno.) Os tokens de texto/marca/accent/status **não mudam** (já verificados AA anteriormente).

- [ ] **Step 2: Trocar `Badge` de fill suave pra contornado**

Substituir `frontend-react/src/ui/Badge.tsx` inteiro por:

```tsx
import type { ReactNode } from "react";
import { TONE_OUTLINE, type Tone } from "./tone";

type Size = "sm" | "md";

interface BadgeProps {
  tone?: Tone;
  size?: Size;
  children: ReactNode;
}

export function Badge({ tone = "neutral", size = "sm", children }: BadgeProps) {
  return (
    <span
      className={`inline-flex items-center gap-1 rounded-pill border font-medium ${TONE_OUTLINE[tone]} ${
        size === "sm" ? "px-2 py-0.5 text-xs" : "px-2.5 py-1 text-sm"
      }`}
    >
      {children}
    </span>
  );
}
```

(`TONE_OUTLINE` só define `border-{cor}/40` + `text-{cor}`, sem largura de borda — por isso o `border` explícito na className, senão a borda fica invisível. Mesmo padrão que `ApiHealthBanner.tsx` já usa com `border-b` + `TONE_OUTLINE`.)

- [ ] **Step 3: Trocar os dois pontos de fill suave em `ValidationReport.tsx`**

No topo do arquivo, adicionar o import:

```tsx
import { TONE_OUTLINE } from "./tone";
```

Trocar:

```tsx
        <div className="rounded-control bg-danger-soft px-3 py-2 text-sm text-danger">{report.general_errors}</div>
```

por:

```tsx
        <div className={`rounded-control border px-3 py-2 text-sm ${TONE_OUTLINE.danger}`}>{report.general_errors}</div>
```

E trocar:

```tsx
              <span key={col} className="rounded-pill bg-warning-soft px-2 py-0.5 text-xs text-warning">
```

por:

```tsx
              <span key={col} className={`rounded-pill border px-2 py-0.5 text-xs ${TONE_OUTLINE.warning}`}>
```

- [ ] **Step 4: Verificar**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test`
Expected: tudo passa (nenhum teste cobre essas classes, então nenhuma mudança de contagem).

- [ ] **Step 5: Checar manualmente (contraste)**

`npm run dev`, abrir em dark: conferir Badge (Home, Estruturas), `ValidationReport` (Cadastro, validar uma planilha com erro) e `ApiHealthBanner` (derrubar a API ou usar DevTools pra simular offline) — texto e borda devem estar legíveis sobre os novos tons de superfície. Repetir em light.

- [ ] **Step 6: Commit**

```bash
git add frontend-react/src/index.css frontend-react/src/ui/Badge.tsx frontend-react/src/ui/ValidationReport.tsx
git commit -m "feat(frontend): paleta dark-first (bordas mais fortes, menos camada de cor) e badges contornados"
```

---

## Task 4: Sombras no dark, transição de navegação e microinterações

**Files:**
- Modify: `frontend-react/src/index.css:48-50` (tokens de sombra em `@theme`)
- Modify: `frontend-react/src/index.css:64-92` e `:96-127` (adicionar `--pd-shadow-*` em `:root` e `.dark`)
- Modify: `frontend-react/src/index.css` (adicionar CSS da transição `vt-nav` no final do arquivo)
- Modify: `frontend-react/src/ui/Card.tsx` (linha do `interactive &&` dentro do `cn(...)`)
- Modify: `frontend-react/src/ui/Button.tsx` (linha de classes base)

Sem teste automatizado — visual. `Modal.tsx` não precisa de nenhuma edição: já usa `shadow-pop`, herda a mudança automaticamente.

- [ ] **Step 1: Indireção dos tokens de sombra**

Em `frontend-react/src/index.css`, dentro do bloco `@theme` (linhas 48-50), trocar:

```css
  --shadow-card: 0 1px 2px rgb(16 24 40 / 0.06), 0 1px 3px rgb(16 24 40 / 0.1);
  --shadow-card-hover: 0 6px 20px rgb(16 24 40 / 0.1);
  --shadow-pop: 0 12px 32px rgb(16 24 40 / 0.18);
```

por:

```css
  --shadow-card: var(--pd-shadow-card);
  --shadow-card-hover: var(--pd-shadow-card-hover);
  --shadow-pop: var(--pd-shadow-pop);
```

(Mesma indireção que `--color-*` já usa pra virar sensível a tema.)

- [ ] **Step 2: Valores em `:root`**

No bloco `:root` (linhas 64-92), adicionar ao final, antes do `}` de fechamento:

```css

  --pd-shadow-card: 0 1px 2px rgb(16 24 40 / 0.06), 0 1px 3px rgb(16 24 40 / 0.1);
  --pd-shadow-card-hover: 0 6px 20px rgb(16 24 40 / 0.1);
  --pd-shadow-pop: 0 12px 32px rgb(16 24 40 / 0.18);
```

(Valores idênticos aos de hoje — light mode não muda.)

- [ ] **Step 3: Valores em `.dark`**

No bloco `.dark`, adicionar ao final, antes do `}` de fechamento:

```css

  --pd-shadow-card: 0 0 0 1px rgb(0 0 0 / 0.3);
  --pd-shadow-card-hover: 0 4px 16px rgb(0 0 0 / 0.4);
  --pd-shadow-pop: 0 16px 40px rgb(0 0 0 / 0.55);
```

(Sombra "card" vira essencialmente nula — quem define o card no dark é a borda mais forte da Task 3. `hover`/`pop` continuam dando elevação real, mas em preto em vez do cinza-azulado de `rgb(16 24 40)`, que quase não aparece sobre fundo escuro.)

- [ ] **Step 4: CSS da transição de navegação**

No final de `frontend-react/src/index.css` (depois do `@keyframes pd-toast-in`), adicionar:

```css

/* Transição de troca de aba (navigate() em routing/useHashRoute.ts) — classe
   `vt-nav` liga/desliga em <html> só durante essa transição, pra não afetar
   o crossfade simples que o ThemeToggle já usa (sem essa classe). */
:root.vt-nav::view-transition-old(root) {
  animation: pd-vt-out 150ms ease-in both;
}

:root.vt-nav::view-transition-new(root) {
  animation: pd-vt-in 150ms ease-out both;
}

@keyframes pd-vt-out {
  to {
    opacity: 0;
    transform: translateY(-6px);
  }
}

@keyframes pd-vt-in {
  from {
    opacity: 0;
    transform: translateY(6px);
  }
}
```

- [ ] **Step 5: Microinterações mais perceptíveis**

Em `frontend-react/src/ui/Card.tsx`, no bloco `interactive &&` (dentro do `cn(...)`), trocar:

```ts
          "cursor-pointer transition hover:-translate-y-0.5 hover:border-border-strong hover:shadow-card-hover",
```

por:

```ts
          "cursor-pointer transition hover:-translate-y-1 hover:border-border-strong hover:shadow-card-hover",
```

Em `frontend-react/src/ui/Button.tsx`, no array `SIZES`/classe base, adicionar um leve scale no clique. Trocar a linha de classes base (a string passada em `cn(...)` antes de `VARIANTS[variant]`):

```ts
        "inline-flex items-center justify-center rounded-control font-medium transition-colors outline-none focus-visible:ring-2 focus-visible:ring-accent/50 disabled:cursor-not-allowed disabled:opacity-50",
```

por:

```ts
        "inline-flex items-center justify-center rounded-control font-medium transition outline-none focus-visible:ring-2 focus-visible:ring-accent/50 active:scale-[0.97] disabled:cursor-not-allowed disabled:opacity-50 disabled:active:scale-100",
```

(`transition-colors` → `transition` pra também animar o `transform` do `active:scale`; `disabled:active:scale-100` evita o efeito em botão desabilitado.)

- [ ] **Step 6: Verificar**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test && npm run build`
Expected: tudo passa, build gera `frontend/static/react/app-react.css` sem erro (confirma que o Tailwind/Lightning CSS aceita `@starting-style`/pseudo-elementos de view-transition sem quebrar o pipeline — checar nesta etapa, antes da Task 5 usar `@starting-style` de verdade).

- [ ] **Step 7: Checar manualmente**

Comparar Card/Modal em dark antes/depois (sombra quase some, borda define a forma) e reconferir a transição de troca de aba da Task 2 (agora com o slide+fade customizado em vez do crossfade padrão do browser). Passar o mouse num card interativo da Home (lift mais perceptível) e clicar num botão (leve "afundada" de escala).

- [ ] **Step 8: Commit**

```bash
git add frontend-react/src/index.css frontend-react/src/ui/Card.tsx frontend-react/src/ui/Button.tsx
git commit -m "feat(frontend): sombra quase nula no dark, transição entre telas e microinterações mais perceptíveis"
```

---

## Task 5: Utilitários de entrada de conteúdo e stagger

**Files:**
- Modify: `frontend-react/src/index.css` (adicionar `@layer utilities` no final do arquivo)

Infra pura — sem uso ainda (Tasks 6/7 aplicam). Sem teste automatizado.

- [ ] **Step 1: Adicionar as classes utilitárias**

No final de `frontend-react/src/index.css` (depois do CSS adicionado na Task 4), adicionar:

```css

/* Entrada de conteúdo (`.pd-enter`) e escalonamento (`.pd-stagger`) —
   usados por Home, Histórico e a tabela de Estruturas. `@starting-style`
   dá o estado inicial da transição no 1º paint/inserção no DOM, sem JS;
   `--i` (setado inline por item, ex. style={{ "--i": index }}) atrasa cada
   item em `40ms * i` via transition-delay. Cobertos pelo gate global de
   prefers-reduced-motion no topo deste arquivo (zera duration/delay). */
@layer utilities {
  .pd-enter {
    opacity: 1;
    transform: translateY(0);
    transition:
      opacity 240ms ease-out,
      transform 240ms ease-out;
  }

  @starting-style {
    .pd-enter {
      opacity: 0;
      transform: translateY(8px);
    }
  }

  .pd-stagger {
    transition-delay: calc(var(--i, 0) * 40ms);
  }
}
```

- [ ] **Step 2: Verificar**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test && npm run build`
Expected: tudo passa.

- [ ] **Step 3: Commit**

```bash
git add frontend-react/src/index.css
git commit -m "feat(frontend): utilitários CSS de entrada de conteúdo (@starting-style) e stagger"
```

---

## Task 6: Aplicar entrada/stagger na Home

**Files:**
- Modify: `frontend-react/src/views/Home/index.tsx`

**Interfaces:**
- Consumes: classes `.pd-enter`/`.pd-stagger` e a custom property `--i` (Task 5).

- [ ] **Step 1: Import de `CSSProperties`**

No topo de `frontend-react/src/views/Home/index.tsx`, ajustar o import de tipo do React:

```tsx
import type { CSSProperties, ReactNode } from "react";
```

- [ ] **Step 2: Cards de operação (`OPS.map`)**

No `.map` de `OPS` (dentro do `<div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">`), adicionar índice e aplicar as classes/estilo no `<Card>`:

```tsx
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
```

- [ ] **Step 3: Lista de "Últimas execuções"**

No `.map` de `recent` (dentro de `<div className="space-y-2">`), envolver cada `<ResultCard>` com um wrapper que carrega a animação (não editar `ResultCard.tsx` — ele não expõe `className`/`style`, e não é usado em nenhum outro lugar sem stagger, então um wrapper aqui é mais simples que crescer a API do componente):

```tsx
        <div className="space-y-2">
          {recent.map((item, index) => (
            <div key={item.key} className="pd-enter pd-stagger" style={{ "--i": index } as CSSProperties}>
              <ResultCard
                operationLabel={item.title}
                timestamp={item.ts}
                inputSummary={item.subtitle ?? ""}
                outputFilename={item.outputFilename}
                status={item.status}
                compact
              />
            </div>
          ))}
        </div>
```

(Note: a prop `key` sai do `<ResultCard>` e vai pro `<div>` wrapper — é ele quem precisa da key agora.)

- [ ] **Step 4: Verificar**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test`
Expected: tudo passa.

- [ ] **Step 5: Checar manualmente**

`npm run dev`, abrir a Home (ou navegar até ela vindo de outra aba) — os 4 cards de operação devem entrar em sequência (leve atraso entre um e outro), e o mesmo pra lista de execuções recentes (gerar algumas execuções primeiro, ou usar o Histórico com dados de teste).

- [ ] **Step 6: Commit**

```bash
git add frontend-react/src/views/Home/index.tsx
git commit -m "feat(frontend): anima entrada dos cards e da lista de execuções recentes na Home"
```

---

## Task 7: Aplicar entrada/stagger no Histórico e na tabela de Estruturas

**Files:**
- Modify: `frontend-react/src/tabs/HistoricoTab/index.tsx` (linhas do `<tbody>`, ~124-133)
- Modify: `frontend-react/src/tabs/EstruturasTab/index.tsx` (linhas do `<tbody>`, ~328-369)

Índice é limitado a 12 (`Math.min(index, 12)`) pra listas grandes não deixarem os últimos itens entrando meio segundo depois dos primeiros — acima de 12 itens, o atraso fica fixo em `12 * 40ms = 480ms`.

- [ ] **Step 1: `HistoricoTab.tsx`**

No import de tipos, adicionar `CSSProperties`:

```tsx
import type { CSSProperties } from "react";
```

(No topo do arquivo, junto dos outros imports.)

No `.map` de `filtered` (dentro do `<tbody id="historico_tbody">`), trocar:

```tsx
              {filtered.map((item, i) => (
                <tr key={`${item.ts}-${i}`} className="border-t border-border">
```

por:

```tsx
              {filtered.map((item, i) => (
                <tr
                  key={`${item.ts}-${i}`}
                  className="pd-enter pd-stagger border-t border-border"
                  style={{ "--i": Math.min(i, 12) } as CSSProperties}
                >
```

- [ ] **Step 2: `EstruturasTab.tsx`**

No import de tipos, adicionar `CSSProperties`:

```tsx
import type { CSSProperties } from "react";
```

No `.map` de `e.items` (dentro de `<tbody id="aprovacao_table_body">`), adicionar o índice e aplicar a animação na `<tr>`:

```tsx
                  {e.items.map((item, index) => {
                    const flagged = isSubstituir ? item.teraDuplicidade : item.ficaraSemAprovador;
                    const flagTitle = isSubstituir
                      ? "O novo aprovador já está presente nesta estrutura"
                      : "Esta estrutura ficará sem aprovadores";
                    return (
                      <tr
                        key={item.aprovacaoId}
                        className={`pd-enter pd-stagger border-t border-border ${flagged ? "bg-warning/15" : ""}`}
                        style={{ "--i": Math.min(index, 12) } as CSSProperties}
                      >
```

(Resto do `<tr>` — `<td>`s internos — fica igual, só a abertura da tag muda.)

- [ ] **Step 3: Verificar**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test`
Expected: tudo passa.

- [ ] **Step 4: Checar manualmente**

Histórico: com itens no histórico (gerar algumas execuções ou usar dados de teste), as linhas devem entrar escalonadas ao abrir a aba. **Atenção**: digitar no campo "Filtrar por texto" refiltra a lista e reindexar pode re-disparar a animação em várias linhas a cada tecla — conferir se isso incomoda visualmente; se sim, é um ajuste de keying (fora do escopo deste plano, anotar como follow-up). Estruturas: com uma base carregada e "Verificar" clicado, as linhas da tabela de estruturas devem entrar escalonadas.

- [ ] **Step 5: Commit**

```bash
git add frontend-react/src/tabs/HistoricoTab/index.tsx frontend-react/src/tabs/EstruturasTab/index.tsx
git commit -m "feat(frontend): anima entrada das linhas no Histórico e na tabela de Estruturas"
```

---

## Task 8: Componente Skeleton + estados de carregamento

**Files:**
- Create: `frontend-react/src/ui/Skeleton.tsx`
- Modify: `frontend-react/src/index.css` (adicionar `@keyframes pd-shimmer`)
- Modify: `frontend-react/src/ui/ValidationReport.tsx` (branch `if (loading)`)
- Modify: `frontend-react/src/tabs/InativacaoTab/index.tsx` (tabela de resultados, dentro de `showResults`)

**Interfaces:**
- Produces: `Skeleton({ className? }: { className?: string })` — `<div>` decorativo (`aria-hidden`), de `frontend-react/src/ui/Skeleton.tsx`, consumido por `ValidationReport.tsx` e `InativacaoTab.tsx`.

- [ ] **Step 1: Keyframe do shimmer**

No final de `frontend-react/src/index.css`, adicionar:

```css

@keyframes pd-shimmer {
  0%,
  100% {
    opacity: 1;
  }
  50% {
    opacity: 0.5;
  }
}
```

- [ ] **Step 2: Criar `Skeleton.tsx`**

```tsx
import { cn } from "./cn";

interface SkeletonProps {
  className?: string;
}

/** Placeholder de carregamento — bloco com pulso via CSS puro
 *  (@keyframes pd-shimmer em index.css). Puramente decorativo: quem usa
 *  este componente continua anunciando o texto real do estado de loading
 *  via aria-live/sr-only, já que este `<div>` é aria-hidden. */
export function Skeleton({ className }: SkeletonProps) {
  return (
    <div
      aria-hidden="true"
      className={cn("rounded-control bg-surface-sunken [animation:pd-shimmer_1.4s_ease-in-out_infinite]", className)}
    />
  );
}
```

- [ ] **Step 3: Usar no loading do `ValidationReport`**

Em `frontend-react/src/ui/ValidationReport.tsx`, adicionar o import:

```tsx
import { Skeleton } from "./Skeleton";
```

Trocar o branch `if (loading)`:

```tsx
  if (loading) {
    return (
      <div className="rounded-surface border border-border bg-surface-2 px-4 py-6 text-center text-sm text-text-muted">
        Validando planilha…
      </div>
    );
  }
```

por:

```tsx
  if (loading) {
    return (
      <div className="rounded-surface border border-border bg-surface-2 p-4" aria-live="polite">
        <span className="sr-only">Validando planilha…</span>
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-4" aria-hidden="true">
          <Skeleton className="h-14" />
          <Skeleton className="h-14" />
          <Skeleton className="h-14" />
          <Skeleton className="h-14" />
        </div>
      </div>
    );
  }
```

(O grid de 4 blocos imita o layout dos `StatCard` que aparecem quando o relatório carrega — ver o restante de `ValidationReport.tsx` logo abaixo desse branch.)

- [ ] **Step 4: Usar no loading da tabela de resultados da Inativação**

Em `frontend-react/src/tabs/InativacaoTab/index.tsx`, adicionar o import:

```tsx
import { Skeleton } from "../../ui/Skeleton";
```

No `<tbody id="results_body">`, trocar:

```tsx
                <tbody id="results_body">
                  {pageRows.map((r, i) => (
```

por:

```tsx
                <tbody id="results_body">
                  {inativacao.searching && pageRows.length === 0 && (
                    <tr aria-hidden="true">
                      <td colSpan={4} className="px-3 py-2">
                        <div className="space-y-1.5">
                          <Skeleton className="h-5 w-full" />
                          <Skeleton className="h-5 w-full" />
                          <Skeleton className="h-5 w-full" />
                        </div>
                      </td>
                    </tr>
                  )}
                  {pageRows.map((r, i) => (
```

(A tabela já fica montada com `showResults` mesmo antes de haver resultados — ver `const showResults = inativacao.results.length > 0 || inativacao.searching;` no topo do componente — então o `<tbody>` hoje fica vazio durante a 1ª busca; o skeleton preenche esse vazio.)

- [ ] **Step 5: Verificar**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test && npm run build`
Expected: tudo passa.

- [ ] **Step 6: Checar manualmente**

Cadastro: selecionar uma planilha e clicar "Validar planilha" — durante o carregamento deve aparecer o grid de 4 blocos pulsando em vez do texto "Validando planilha…" (texto continua existindo pra leitor de tela, só visualmente escondido). Inativação: colar alguns CPFs/nomes e clicar "Buscar Usuários" — durante a busca, 3 linhas pulsando devem aparecer na tabela antes dos resultados reais.

- [ ] **Step 7: Commit**

```bash
git add frontend-react/src/ui/Skeleton.tsx frontend-react/src/index.css frontend-react/src/ui/ValidationReport.tsx frontend-react/src/tabs/InativacaoTab/index.tsx
git commit -m "feat(frontend): skeleton de carregamento na validação de planilha e na busca de inativação"
```

---

## Task 9: Verificação final

- [ ] **Step 1: Suíte completa**

Run: `cd frontend-react && npx tsc -b --noEmit && npm run lint && npm test && npm run build`
Expected: tudo passa (tsc limpo, lint limpo, todos os testes — 32 esperados — e build de biblioteca gerando `frontend/static/react/app-react.{js,css}` sem erro).

- [ ] **Step 2: Checklist visual manual (light + dark, via `npm run dev`)**

- [ ] Recarregar sem `localStorage` salvo → abre em dark (Task 1).
- [ ] Alternar tema pelo `ThemeToggle` → crossfade continua funcionando (sem regressão do que já existia).
- [ ] Navegar entre Início/Cadastro/Inativação/Estruturas/Histórico → transição sutil de slide+fade (Task 2/4).
- [ ] Cards da Home, execuções recentes, linhas do Histórico e da tabela de Estruturas → entram escalonados (Task 5/6/7).
- [ ] Badge (Home, Estruturas) e avisos do `ValidationReport` → contornados, texto legível em dark e light (Task 3).
- [ ] Card/Modal em dark → sombra quase imperceptível, forma definida pela borda (Task 4).
- [ ] Validar planilha no Cadastro / buscar na Inativação → skeleton pulsando durante o carregamento (Task 8).
- [ ] DevTools → "Emulate CSS prefers-reduced-motion: reduce" → todas as animações acima (crossfade, transição de tela, entrada/stagger, skeleton) somem ou viram instantâneas.
- [ ] Conferir contraste (DevTools color picker ou similar) do texto sobre os novos tons de superfície do dark — mínimo AA (4.5:1 pra texto normal, 3:1 pra texto grande/ícones).

- [ ] **Step 3: Nada para commitar nesta task** (é só verificação — se algo falhar, voltar pra task correspondente, corrigir, e recommitar lá).
