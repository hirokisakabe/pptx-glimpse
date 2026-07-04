import { afterEach, describe, expect, it } from "vitest";

import {
  _resetCjkFallbackCache,
  _setCjkFallbackPlatformForTest,
  getCjkFallbackFonts,
} from "./cjk-font-fallback.js";

describe("getCjkFallbackFonts", () => {
  afterEach(() => {
    _resetCjkFallbackCache();
  });

  it("On macOS, returns Gothic style such as Hiragino Sans.", () => {
    _setCjkFallbackPlatformForTest("darwin");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("macOS returns Hiragino Mincho ProN to Mincho series", () => {
    _setCjkFallbackPlatformForTest("darwin");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Hiragino Mincho ProN"]);
  });

  it("Returns the same fallback for Noto Sans CJK JP on macOS", () => {
    _setCjkFallbackPlatformForTest("darwin");
    const result = getCjkFallbackFonts("Noto Sans CJK JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("On Windows, return Yu Gothic etc. for Gothic type.", () => {
    _setCjkFallbackPlatformForTest("win32");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Yu Gothic", "Meiryo", "MS Gothic"]);
  });

  it("On Windows, return Yu Mincho etc. for Mincho series.", () => {
    _setCjkFallbackPlatformForTest("win32");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Yu Mincho", "MS Mincho"]);
  });

  it("Returns an empty array on Linux", () => {
    _setCjkFallbackPlatformForTest("linux");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual([]);
  });

  it("Returns an empty array for unknown mapping destinations", () => {
    _setCjkFallbackPlatformForTest("darwin");
    const result = getCjkFallbackFonts("Unknown Font");
    expect(result).toEqual([]);
  });
});
