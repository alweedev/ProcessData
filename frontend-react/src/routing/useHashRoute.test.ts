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
