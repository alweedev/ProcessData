# Redesign dark-first + animações — design

## Contexto

O frontend (`frontend-react/`) já passou por algumas rodadas de polimento nesta
mesma iniciativa: unificação de tokens de radius, `IconChip`/`tone.ts`
centralizando cor, `ApiHealthBanner`, largura de container, grid da Home. O
pedido agora é mais amplo: "fique mais modernizado e adicione mais animações".

Depois de perguntas de escopo, a direção definida foi:

- **Estilo**: dark-first, alto contraste, minimalista — referência Vercel /
  GitHub dark / Supabase. Dark passa a ser o tema padrão na primeira visita;
  light continua existindo (redesenhado com a mesma lógica), acessível pelo
  `ThemeToggle` já existente.
- **Animação**: intensidade "mais pronunciada" (não só microinterações sutis),
  cobrindo 4 áreas: microinterações, entrada de conteúdo, transição entre
  telas, estados de carregamento.
- **Tecnologia**: o usuário perguntou se valia trocar o CSS por outro
  framework "pra usar tecnologia atual". Avaliado e descartado — Tailwind v4
  (motor 100% CSS nativo, sem config JS) + custom properties já É tecnologia
  atual; CSS-in-JS runtime (styled-components/Emotion) seria um passo atrás
  (custo de runtime, comunidade React migrando pra longe disso), e uma opção
  zero-runtime (vanilla-extract/Panda/StyleX) exigiria reescrever a forma de
  estilizar todo componente sem benefício claro pro tamanho do app.
  **Decisão**: manter Tailwind v4 + CSS nativo, e cobrir as 4 áreas de
  animação com features nativas recentes da plataforma (`@starting-style`,
  stagger via custom property, View Transitions API estendida) em vez de
  adicionar uma lib JS de animação (`motion`/Framer Motion) — zero
  dependência nova, e tecnicamente mais "atual" que uma lib JS de 2018.
- **Público-alvo**: ferramenta interna mirando Windows 11/Edge-Chrome (o
  próprio `index.css` já assume isso no font-stack). `@starting-style` e
  View Transitions em navegação degradam bem em navegadores sem suporte
  (elemento só aparece sem animar, sem erro) — aceito como trade-off.

## Fora do escopo

- Trocar Tailwind por outro framework/lib de CSS.
- Adicionar biblioteca de animação JS (Motion/Framer Motion, react-spring,
  etc.).
- Gestos (drag/swipe/pan) — não fazem sentido nesse tipo de ferramenta
  (formulários, tabelas).
- Remover o light mode ou o `ThemeToggle`.
- Mexer no esquema `?v=N` de cache-busting do `frontend/index.html` (não
  relacionado).

## Direção visual (tokens)

Arquivo: `frontend-react/src/index.css`.

Hoje o dark mode já é bem escuro (`--pd-bg: #0a0d12`), mas os tokens `-soft`
(fundo colorido suave em Badge/IconChip/alertas) dão uma leitura mais "SaaS
amigável" que "dark-first minimalista". Ajustes:

- Fundos/superfícies (`--pd-bg`/`-surface`/`-surface-2`/`-surface-sunken`)
  com menos salto de matiz entre camadas; bordas (`--pd-border-strong`)
  ganham mais peso pra definir cards por contorno, já que sombra quase some
  no dark (como GitHub dark) — `--shadow-card`/`-card-hover` reduzidas no
  bloco `.dark`.
- Estados de tom (Badge, `ValidationReport`, `ApiHealthBanner`) migram de
  **fill colorido suave para contornado** (`border-{tone}/40` + fundo neutro,
  em vez de `bg-{tone}-soft`) — reduz ruído de cor, já existe precedente
  disso em `TONE_OUTLINE` (`ui/tone.ts`), que hoje só é usado pelo
  `ApiHealthBanner`; passa a ser o padrão para os demais.
- Accent único usado com intenção: brand (teal) fica reservado pra
  identidade (logo, foco, links); accent (azul) continua sendo a cor de ação
  primária — sem introduzir uma terceira cor competindo.
- Valores hex exatos são ajustados empiricamente durante a implementação,
  conferindo contraste AA no browser (mesmo processo já usado pro fix de
  contraste do `--pd-accent` dark existente) — não travar hex specs aqui sem
  ver renderizado.
- Light mode: mesma lógica (contorno > fill), mas não é mais o tema default
  na primeira visita — vira alternativa igualmente cuidada.
- Tema padrão: o script anti-FOUC e o `ThemeToggle`
  (`frontend-react/src/chrome/ThemeToggle.tsx`) hoje decidem o tema inicial
  a partir de `localStorage`/`prefers-color-scheme`; passa a haver um
  terceiro critério (ausência de preferência salva E sem
  `prefers-color-scheme: light` explícito → default dark), documentado no
  próprio código.

## Componentes

Ajuste visual nos primitivos compartilhados — sem mudar a API pública de
nenhum, só o miolo de classes/tokens (herda pra toda tela automaticamente):

- `ui/Button.tsx`, `ui/Card.tsx`, `ui/Badge.tsx`, `ui/IconChip.tsx` — tratamento
  contornado/alto-contraste.
- `components/Modal.tsx`, `toast/ToastViewport.tsx` — mesma linguagem visual.
- `chrome/ThemeToggle.tsx` — pequenos ajustes pra combinar com o novo
  contraste; a lógica de View Transitions já existente é reaproveitada (ver
  abaixo).

## Sistema de animação (100% CSS/nativo, sem lib nova)

### 1. Microinterações
Continuam em CSS (como já são hoje em `Button`/`Card`/nav) — refinadas com
leve scale/lift mais perceptível (`transform` + `transition`), já cobertas
pelo gate `@media (prefers-reduced-motion: reduce)` existente em
`index.css`.

### 2. Entrada de conteúdo
`@starting-style` (novo bloco `@layer base` em `index.css`, ou classe
utilitária reaproveitável) define o estado inicial (opacidade/translateY) de
elementos que entram no DOM — cards da Home, `ResultCard`/linhas do
Histórico, linhas da tabela de `EstruturasTab`. Sem JS: o navegador anima
sozinho do "starting style" pro estado final assim que o elemento é
inserido/torna-se visível.

### 3. Stagger em listas
Custom property `--i` setada inline por item (`style={{ "--i": index }}` nos
`.map()` já existentes de `OPS`/`recent`/linhas de tabela) +
`animation-delay: calc(var(--i) * 40ms)` numa classe utilitária — escalona a
entrada sem nenhuma lib, funciona em qualquer navegador (é só
`animation-delay`, universal).

### 4. Transição entre telas
Estende o padrão já usado em `ThemeToggle.tsx`
(`document.startViewTransition`) pra troca de aba em
`routing/useHashRoute.ts`/`App.tsx`: a navegação entre `home`/`cadastro`/
`inativacao`/`estruturas`/`historico` passa a ser envolvida por
`startViewTransition` quando disponível (mesmo guard de
`prefers-reduced-motion` já usado no toggle), com uma transição customizada
via `::view-transition-old(root)`/`::view-transition-new(root)` no CSS (algo
como slide+fade curto, mais sutil que o crossfade de tema).

### 5. Estados de carregamento
Novo `ui/Skeleton.tsx` — barras/blocos com `@keyframes` de shimmer/pulse em
CSS puro, substituindo os textos estáticos "Validando...", "Buscando...",
"Processando..." nos pontos de loading de `CadastroTab`, `InativacaoTab`,
`EstruturasTab` (mantendo o texto para leitores de tela via
`aria-live`/`sr-only`, só a representação visual vira skeleton).

### Acessibilidade
Tudo (CSS transitions, `@starting-style`, stagger, View Transitions,
skeleton) fica coberto pelo único gate já existente,
`@media (prefers-reduced-motion: reduce)` em `index.css` — não é necessário
nenhum mecanismo paralelo (diferente do que aconteceria se tivéssemos
adicionado uma lib JS com seu próprio sistema de reduced-motion).

## Rollout

1. Tokens (`index.css`) + componentes base (`ui/*`, `components/Modal`,
   `toast/ToastViewport`) — todas as telas herdam automaticamente.
2. Infra de animação: bloco `@starting-style`/stagger utilitário em
   `index.css`, `ui/Skeleton.tsx`, extensão do `startViewTransition` em
   `App.tsx`/`useHashRoute.ts`.
3. Home (tela mais simples, valida o padrão de entrada/stagger).
4. Cadastro → Inativação → Estruturas (loading states + stagger de tabela).
5. Histórico (tabela mais densa) → Modal/Toast por último.

## Testes / verificação

- `tsc -b --noEmit`, `oxlint`, `vitest run` (`frontend-react/`) — continuam
  valendo, sem mudança de comportamento testável nos hooks/lógica existente.
- `npm run build` — confirma que o build de biblioteca (Vite `lib` mode)
  gera CSS/JS válido com as novas regras.
- Verificação visual manual via browser (light + dark) em cada etapa do
  rollout: contraste AA nas novas cores, animações disparando corretamente,
  e confirmação de que `prefers-reduced-motion` desliga tudo (emular via
  DevTools/`Emulate CSS media feature`).
- Sem teste automatizado novo para as animações em si (comportamento visual,
  não lógico) — cobertura via inspeção manual, como já vem sendo feito nesta
  sessão.
