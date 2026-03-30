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

  it("macOS ではゴシック系に Hiragino Sans 等を返す", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("macOS では明朝系に Hiragino Mincho ProN を返す", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Hiragino Mincho ProN"]);
  });

  it("macOS では Noto Sans CJK JP にも同じフォールバックを返す", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Sans CJK JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("Windows ではゴシック系に Yu Gothic 等を返す", () => {
    mockPlatform.mockReturnValue("win32");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Yu Gothic", "Meiryo", "MS Gothic"]);
  });

  it("Windows では明朝系に Yu Mincho 等を返す", () => {
    mockPlatform.mockReturnValue("win32");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Yu Mincho", "MS Mincho"]);
  });

  it("Linux では空配列を返す", () => {
    mockPlatform.mockReturnValue("linux");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual([]);
  });

  it("未知のマッピング先では空配列を返す", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Unknown Font");
    expect(result).toEqual([]);
  });
});
