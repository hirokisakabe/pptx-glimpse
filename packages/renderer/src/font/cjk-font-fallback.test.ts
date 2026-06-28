import { afterEach, describe, expect, it, vi } from "vitest";

import { _resetCjkFallbackCache, getCjkFallbackFonts } from "./cjk-font-fallback.js";

vi.mock("node:os", () => ({
  platform: vi.fn(),
}));

import { platform } from "node:os";

const mockPlatform = vi.mocked(platform);

describe("getCjkFallbackFonts", () => {
  afterEach(() => {
    _resetCjkFallbackCache();
  });

  it("covers cjk-font-fallback behavior 1", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("covers cjk-font-fallback behavior 2", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Hiragino Mincho ProN"]);
  });

  it("covers cjk-font-fallback behavior 3", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Sans CJK JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("covers cjk-font-fallback behavior 4", () => {
    mockPlatform.mockReturnValue("win32");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Yu Gothic", "Meiryo", "MS Gothic"]);
  });

  it("covers cjk-font-fallback behavior 5", () => {
    mockPlatform.mockReturnValue("win32");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Yu Mincho", "MS Mincho"]);
  });

  it("covers cjk-font-fallback behavior 6", () => {
    mockPlatform.mockReturnValue("linux");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual([]);
  });

  it("covers cjk-font-fallback behavior 7", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Unknown Font");
    expect(result).toEqual([]);
  });
});
