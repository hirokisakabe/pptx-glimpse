import { afterEach, describe, expect, it, vi } from "vitest";

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
    getAdvanceWidth: (text: string, fontSize: number) => text.length * fontSize * 0.6,
  };
}

describe("TextPathFontResolver context", () => {
  afterEach(() => {
    resetTextPathFontResolver();
  });

  it("Returns null in the initial state", () => {
    expect(getTextPathFontResolver()).toBeNull();
  });

  it("Can be obtained using get after set", () => {
    const fonts = new Map([["Arial", createMockFont("Arial")]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    setTextPathFontResolver(resolver);
    expect(getTextPathFontResolver()).toBe(resolver);
  });

  it("Returns to null after reset", () => {
    const fonts = new Map([["Arial", createMockFont("Arial")]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    setTextPathFontResolver(resolver);
    resetTextPathFontResolver();
    expect(getTextPathFontResolver()).toBeNull();
  });
});

describe("DefaultTextPathFontResolver", () => {
  it("Resolve fonts with fontFamily", () => {
    const arialFont = createMockFont("Arial");
    const fonts = new Map([["Arial", arialFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Arial", null)).toBe(arialFont);
  });

  it("If fontFamily is not found, resolve with fontFamilyEa", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const fonts = new Map([["Noto Sans JP", notoFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Unknown", "Noto Sans JP")).toBe(notoFont);
  });

  it("If neither is found, fallback to defaultFont", () => {
    const defaultFont = createMockFont("Default");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts, defaultFont);
    expect(resolver.resolveFont("Unknown", "AlsoUnknown")).toBe(defaultFont);
  });

  it("If there is none, return null", () => {
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Unknown", null)).toBeNull();
  });

  it("If fontFamily is null, try fontFamilyEa", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const fonts = new Map([["Noto Sans JP", notoFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont(null, "Noto Sans JP")).toBe(notoFont);
  });
});

describe("DefaultTextPathFontResolver CJK fallback", () => {
  afterEach(() => {
    resetFontMapping();
    vi.restoreAllMocks();
  });

  it("Attempt CJK fallback chain if no mapping is found", async () => {
    // In Linux CI, the fallback chain is empty, so mock the macOS equivalent value.
    const mod = await import("./cjk-font-fallback.js");
    vi.spyOn(mod, "getCjkFallbackFonts").mockReturnValue([
      "Hiragino Sans",
      "Hiragino Kaku Gothic ProN",
    ]);

    const hiraginoFont = createMockFont("Hiragino Sans");
    const fonts = new Map([["Hiragino Sans", hiraginoFont]]);
    setFontMapping({ Meiryo: "Noto Sans JP" });
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Meiryo", null)).toBe(hiraginoFont);
  });

  it("Returns the second font in the CJK fallback chain", async () => {
    // In Linux CI, the fallback chain is empty, so mock the macOS equivalent value.
    const mod = await import("./cjk-font-fallback.js");
    vi.spyOn(mod, "getCjkFallbackFonts").mockReturnValue([
      "Hiragino Sans",
      "Hiragino Kaku Gothic ProN",
    ]);

    const kakuFont = createMockFont("Hiragino Kaku Gothic ProN");
    const fonts = new Map([["Hiragino Kaku Gothic ProN", kakuFont]]);
    setFontMapping({ Meiryo: "Noto Sans JP" });
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Meiryo", null)).toBe(kakuFont);
  });
});

describe("DefaultTextPathFontResolver font warnings", () => {
  afterEach(() => {
    resetFontMapping();
    initWarningLogger("off");
  });

  it("Issue font.notFound warning if font is not found", () => {
    initWarningLogger("warn");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("UnknownFont", null);
    const entries = getWarningEntries();
    expect(entries.some((e) => e.feature === "font.notFound")).toBe(true);
    expect(entries.some((e) => e.message.includes("UnknownFont"))).toBe(true);
  });

  it("Warnings with the same font name are not duplicated", () => {
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

  it("Don't warn if font is found", () => {
    initWarningLogger("warn");
    const arialFont = createMockFont("Arial");
    const fonts = new Map([["Arial", arialFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("Arial", null);
    const entries = getWarningEntries().filter((e) => e.feature === "font.notFound");
    expect(entries).toHaveLength(0);
  });
});
