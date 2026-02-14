import { describe, it, expect, afterEach } from "vitest";
import {
  DefaultTextPathFontResolver,
  setTextPathFontResolver,
  getTextPathFontResolver,
  resetTextPathFontResolver,
} from "./text-path-context.js";
import type { OpentypeFullFont } from "./text-path-context.js";

function createMockFont(name: string): OpentypeFullFont {
  return {
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    getPath: (_text: string, _x: number, _y: number, _fontSize: number) => ({
      toPathData: () => `M0 0 L10 10 /* ${name} */`,
    }),
  };
}

describe("TextPathFontResolver context", () => {
  afterEach(() => {
    resetTextPathFontResolver();
  });

  it("初期状態では null を返す", () => {
    expect(getTextPathFontResolver()).toBeNull();
  });

  it("set 後に get で取得できる", () => {
    const fonts = new Map([["Arial", createMockFont("Arial")]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    setTextPathFontResolver(resolver);
    expect(getTextPathFontResolver()).toBe(resolver);
  });

  it("reset 後に null に戻る", () => {
    const fonts = new Map([["Arial", createMockFont("Arial")]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    setTextPathFontResolver(resolver);
    resetTextPathFontResolver();
    expect(getTextPathFontResolver()).toBeNull();
  });
});

describe("DefaultTextPathFontResolver", () => {
  it("fontFamily でフォントを解決する", () => {
    const arialFont = createMockFont("Arial");
    const fonts = new Map([["Arial", arialFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Arial", null)).toBe(arialFont);
  });

  it("fontFamily が見つからなければ fontFamilyEa で解決する", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const fonts = new Map([["Noto Sans JP", notoFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Unknown", "Noto Sans JP")).toBe(notoFont);
  });

  it("どちらも見つからなければ defaultFont にフォールバック", () => {
    const defaultFont = createMockFont("Default");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts, defaultFont);
    expect(resolver.resolveFont("Unknown", "AlsoUnknown")).toBe(defaultFont);
  });

  it("どれもなければ null を返す", () => {
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Unknown", null)).toBeNull();
  });

  it("fontFamily が null の場合は fontFamilyEa を試みる", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const fonts = new Map([["Noto Sans JP", notoFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont(null, "Noto Sans JP")).toBe(notoFont);
  });
});
