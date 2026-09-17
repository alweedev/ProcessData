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
