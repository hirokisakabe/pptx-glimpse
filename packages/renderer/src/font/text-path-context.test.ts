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

  it("covers text-path-context behavior 1", () => {
    expect(getTextPathFontResolver()).toBeNull();
  });

  it("covers text-path-context behavior 2", () => {
    const fonts = new Map([["Arial", createMockFont("Arial")]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    setTextPathFontResolver(resolver);
    expect(getTextPathFontResolver()).toBe(resolver);
  });

  it("covers text-path-context behavior 3", () => {
    const fonts = new Map([["Arial", createMockFont("Arial")]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    setTextPathFontResolver(resolver);
    resetTextPathFontResolver();
    expect(getTextPathFontResolver()).toBeNull();
  });
});

describe("DefaultTextPathFontResolver", () => {
  it("covers text-path-context behavior 4", () => {
    const arialFont = createMockFont("Arial");
    const fonts = new Map([["Arial", arialFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Arial", null)).toBe(arialFont);
  });

  it("covers text-path-context behavior 5", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const fonts = new Map([["Noto Sans JP", notoFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Unknown", "Noto Sans JP")).toBe(notoFont);
  });

  it("covers text-path-context behavior 6", () => {
    const defaultFont = createMockFont("Default");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts, defaultFont);
    expect(resolver.resolveFont("Unknown", "AlsoUnknown")).toBe(defaultFont);
  });

  it("covers text-path-context behavior 7", () => {
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont("Unknown", null)).toBeNull();
  });

  it("covers text-path-context behavior 8", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const fonts = new Map([["Noto Sans JP", notoFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    expect(resolver.resolveFont(null, "Noto Sans JP")).toBe(notoFont);
  });
});

describe("font/text-path-context.test behavior", () => {
  afterEach(() => {
    resetFontMapping();
    vi.restoreAllMocks();
  });

  it("covers text-path-context behavior 9", async () => {
    // Test note.
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

  it("covers text-path-context behavior 10", async () => {
    // Test note.
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

describe("font/text-path-context.test behavior", () => {
  afterEach(() => {
    resetFontMapping();
    initWarningLogger("off");
  });

  it("covers text-path-context behavior 11", () => {
    initWarningLogger("warn");
    const fonts = new Map<string, OpentypeFullFont>();
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("UnknownFont", null);
    const entries = getWarningEntries();
    expect(entries.some((e) => e.feature === "font.notFound")).toBe(true);
    expect(entries.some((e) => e.message.includes("UnknownFont"))).toBe(true);
  });

  it("covers text-path-context behavior 12", () => {
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

  it("covers text-path-context behavior 13", () => {
    initWarningLogger("warn");
    const arialFont = createMockFont("Arial");
    const fonts = new Map([["Arial", arialFont]]);
    const resolver = new DefaultTextPathFontResolver(fonts);
    resolver.resolveFont("Arial", null);
    const entries = getWarningEntries().filter((e) => e.feature === "font.notFound");
    expect(entries).toHaveLength(0);
  });
});
