import { afterEach, describe, expect, it } from "vitest";

import { getWarningEntries, initWarningLogger } from "../warning-logger.js";
import { resetFontMapping, setFontMapping } from "./font-mapping-context.js";
import type { OpentypeFullFont } from "./text-path-context.js";
import {
  DefaultTextPathFontResolver,
  getTextPathFontResolver,
  resetTextPathFontResolver,
  setTextPathFontResolver,
} from "./text-path-context.js";

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

describe("DefaultTextPathFontResolver CJK フォールバック", () => {
  afterEach(() => {
    resetFontMapping();
  });

  it("マッピング先が見つからない場合に CJK フォールバックチェーンを試行する", () => {
    const hiraginoFont = createMockFont("Hiragino Sans");
    const fonts = new Map([["Hiragino Sans", hiraginoFont]]);
    // Meiryo → Noto Sans JP (マッピング) → 見つからない → Hiragino Sans (フォールバック)
    setFontMapping({ Meiryo: "Noto Sans JP" });
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Meiryo", null)).toBe(hiraginoFont);
  });

  it("CJK フォールバックチェーンの2番目のフォントを返す", () => {
    const kakuFont = createMockFont("Hiragino Kaku Gothic ProN");
    const fonts = new Map([["Hiragino Kaku Gothic ProN", kakuFont]]);
    setFontMapping({ Meiryo: "Noto Sans JP" });
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Meiryo", null)).toBe(kakuFont);
  });
});

describe("DefaultTextPathFontResolver font.notFound 警告", () => {
  afterEach(() => {
    resetFontMapping();
    initWarningLogger("off");
  });

  it("フォントが見つからない場合に font.notFound 警告を出す", () => {
    initWarningLogger("warn");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("UnknownFont", null);
    const entries = getWarningEntries();
    expect(entries.some((e) => e.feature === "font.notFound")).toBe(true);
    expect(entries.some((e) => e.message.includes("UnknownFont"))).toBe(true);
  });

  it("同じフォント名の警告は重複しない", () => {
    initWarningLogger("warn");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("UnknownFont", null);
    resolver.resolveFont("UnknownFont", null);
    const entries = getWarningEntries().filter(
      (e) => e.feature === "font.notFound" && e.message.includes("UnknownFont"),
    );
    expect(entries).toHaveLength(1);
  });

  it("フォントが見つかった場合は警告を出さない", () => {
    initWarningLogger("warn");
    const arialFont = createMockFont("Arial");
    const fonts = new Map([["Arial", arialFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("Arial", null);
    const entries = getWarningEntries().filter((e) => e.feature === "font.notFound");
    expect(entries).toHaveLength(0);
  });
});
