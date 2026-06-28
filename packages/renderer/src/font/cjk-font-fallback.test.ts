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

  it("On macOS, returns Gothic style such as Hiragino Sans.", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("macOS returns Hiragino Mincho ProN to Mincho series", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Hiragino Mincho ProN"]);
  });

  it("Returns the same fallback for Noto Sans CJK JP on macOS", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Noto Sans CJK JP");
    expect(result).toEqual(["Hiragino Sans", "Hiragino Kaku Gothic ProN"]);
  });

  it("On Windows, return Yu Gothic etc. for Gothic type.", () => {
    mockPlatform.mockReturnValue("win32");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual(["Yu Gothic", "Meiryo", "MS Gothic"]);
  });

  it("On Windows, return Yu Mincho etc. for Mincho series.", () => {
    mockPlatform.mockReturnValue("win32");
    const result = getCjkFallbackFonts("Noto Serif CJK JP");
    expect(result).toEqual(["Yu Mincho", "MS Mincho"]);
  });

  it("Returns an empty array on Linux", () => {
    mockPlatform.mockReturnValue("linux");
    const result = getCjkFallbackFonts("Noto Sans JP");
    expect(result).toEqual([]);
  });

  it("Returns an empty array for unknown mapping destinations", () => {
    mockPlatform.mockReturnValue("darwin");
    const result = getCjkFallbackFonts("Unknown Font");
    expect(result).toEqual([]);
  });
});
