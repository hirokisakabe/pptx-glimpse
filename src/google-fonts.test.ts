import { describe, it, expect, vi } from "vitest";
import type { UsedFonts } from "./font-collector.js";
import { resolveGoogleFontNames, parseFontUrlsFromCss, fetchGoogleFonts } from "./google-fonts.js";

const sampleUsedFonts: UsedFonts = {
  theme: {
    majorFont: "Calibri Light",
    minorFont: "Calibri",
    majorFontEa: "MS PGothic",
    minorFontEa: "MS PGothic",
    majorFontCs: "Times New Roman",
    minorFontCs: "Arial",
  },
  fonts: ["Arial", "Calibri", "Calibri Light", "MS PGothic", "Times New Roman", "Yu Gothic"],
};

describe("resolveGoogleFontNames", () => {
  it("フォントマッピングを適用して Google Fonts 名を返す", () => {
    const result = resolveGoogleFontNames(sampleUsedFonts);
    expect(result).toContain("Arimo"); // Arial → Arimo
    expect(result).toContain("Carlito"); // Calibri → Carlito
    expect(result).toContain("Tinos"); // Times New Roman → Tinos
    expect(result).toContain("Noto Sans JP"); // MS PGothic, Yu Gothic → Noto Sans JP
  });

  it("重複を除去してソート済みで返す", () => {
    const result = resolveGoogleFontNames(sampleUsedFonts);
    // MS PGothic と Yu Gothic が両方 Noto Sans JP にマッピングされるが重複なし
    const notoCount = result.filter((f) => f === "Noto Sans JP").length;
    expect(notoCount).toBe(1);
    // ソート済み
    const sorted = [...result].sort();
    expect(result).toEqual(sorted);
  });

  it("マッピングに存在しないフォントは除外される", () => {
    const usedFonts: UsedFonts = {
      theme: {
        majorFont: "Unknown Font",
        minorFont: "Another Unknown",
        majorFontEa: null,
        minorFontEa: null,
        majorFontCs: null,
        minorFontCs: null,
      },
      fonts: ["Unknown Font", "Another Unknown"],
    };
    const result = resolveGoogleFontNames(usedFonts);
    expect(result).toEqual([]);
  });

  it("カスタムフォントマッピングを適用できる", () => {
    const usedFonts: UsedFonts = {
      theme: {
        majorFont: "Custom Font",
        minorFont: "Calibri",
        majorFontEa: null,
        minorFontEa: null,
        majorFontCs: null,
        minorFontCs: null,
      },
      fonts: ["Custom Font", "Calibri"],
    };
    const result = resolveGoogleFontNames(usedFonts, {
      "Custom Font": "Roboto",
    });
    expect(result).toContain("Roboto");
    expect(result).toContain("Carlito"); // デフォルトマッピングも維持
  });
});

describe("parseFontUrlsFromCss", () => {
  it("@font-face から font-family と url を抽出する", () => {
    const css = `
/* latin */
@font-face {
  font-family: 'Carlito';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/s/carlito/v3/3Jn9SDPw3m-pk039PDA.ttf) format('truetype');
}
/* latin */
@font-face {
  font-family: 'Arimo';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/s/arimo/v28/P5sfzZCDf9_T_10c.ttf) format('truetype');
}`;
    const result = parseFontUrlsFromCss(css);
    expect(result).toHaveLength(2);
    expect(result[0]).toEqual({
      name: "Carlito",
      url: "https://fonts.gstatic.com/s/carlito/v3/3Jn9SDPw3m-pk039PDA.ttf",
    });
    expect(result[1]).toEqual({
      name: "Arimo",
      url: "https://fonts.gstatic.com/s/arimo/v28/P5sfzZCDf9_T_10c.ttf",
    });
  });

  it("空の CSS では空配列を返す", () => {
    expect(parseFontUrlsFromCss("")).toEqual([]);
  });

  it("url のない @font-face はスキップする", () => {
    const css = `@font-face { font-family: 'Test'; font-style: normal; }`;
    expect(parseFontUrlsFromCss(css)).toEqual([]);
  });

  it("font-family のない @font-face はスキップする", () => {
    const css = `@font-face { src: url(https://example.com/font.ttf); }`;
    expect(parseFontUrlsFromCss(css)).toEqual([]);
  });
});

describe("fetchGoogleFonts", () => {
  const fakeCss = `
@font-face {
  font-family: 'Carlito';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/s/carlito/v3/font.ttf) format('truetype');
}
@font-face {
  font-family: 'Arimo';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/s/arimo/v28/font.ttf) format('truetype');
}`;

  const fakeFont = new Uint8Array([0, 1, 2, 3]);

  function createMockFetch(cssResponse: string, fontData: Uint8Array) {
    return vi.fn((url: string | URL | Request) => {
      const urlStr = typeof url === "string" ? url : url instanceof URL ? url.toString() : url.url;
      if (urlStr.includes("fonts.googleapis.com/css2")) {
        return Promise.resolve(new Response(cssResponse, { status: 200 }));
      }
      if (urlStr.includes("fonts.gstatic.com")) {
        return Promise.resolve(new Response(fontData, { status: 200 }));
      }
      return Promise.resolve(new Response("Not found", { status: 404 }));
    }) as unknown as typeof globalThis.fetch;
  }

  it("Google Fonts からフォントバッファを取得する", async () => {
    const mockFetch = createMockFetch(fakeCss, fakeFont);
    const result = await fetchGoogleFonts(sampleUsedFonts, { fetch: mockFetch });

    expect(result.length).toBeGreaterThan(0);
    expect(result[0]).toHaveProperty("name");
    expect(result[0]).toHaveProperty("data");
    expect(result[0].data).toBeInstanceOf(Uint8Array);
  });

  it("CSS API の URL に必要なフォント名が含まれる", async () => {
    const mockFetch = createMockFetch(fakeCss, fakeFont);
    await fetchGoogleFonts(sampleUsedFonts, { fetch: mockFetch });

    const cssCall = (mockFetch as ReturnType<typeof vi.fn>).mock.calls[0];
    const cssUrl = cssCall[0] as string;
    expect(cssUrl).toContain("fonts.googleapis.com/css2");
    expect(cssUrl).toContain("Arimo");
    expect(cssUrl).toContain("Carlito");
  });

  it("マッピングに該当するフォントがない場合は空配列を返す", async () => {
    const mockFetch = vi.fn() as unknown as typeof globalThis.fetch;
    const usedFonts: UsedFonts = {
      theme: {
        majorFont: "Unknown",
        minorFont: "Unknown2",
        majorFontEa: null,
        minorFontEa: null,
        majorFontCs: null,
        minorFontCs: null,
      },
      fonts: ["Unknown", "Unknown2"],
    };
    const result = await fetchGoogleFonts(usedFonts, { fetch: mockFetch });
    expect(result).toEqual([]);
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it("CSS API がエラーを返した場合は空配列を返す", async () => {
    const mockFetch = vi.fn(() => {
      return Promise.resolve(new Response("Bad Request", { status: 400 }));
    }) as unknown as typeof globalThis.fetch;
    const result = await fetchGoogleFonts(sampleUsedFonts, { fetch: mockFetch });
    expect(result).toEqual([]);
  });

  it("フォントファイル取得のエラーはスキップし、他のフォントは返す", async () => {
    let callCount = 0;
    const mockFetch = vi.fn((url: string | URL | Request) => {
      const urlStr = typeof url === "string" ? url : url instanceof URL ? url.toString() : url.url;
      if (urlStr.includes("fonts.googleapis.com/css2")) {
        return Promise.resolve(new Response(fakeCss, { status: 200 }));
      }
      callCount++;
      if (callCount === 1) {
        return Promise.reject(new Error("Network error"));
      }
      return Promise.resolve(new Response(fakeFont, { status: 200 }));
    }) as unknown as typeof globalThis.fetch;

    const result = await fetchGoogleFonts(sampleUsedFonts, { fetch: mockFetch });
    // 1つはエラーでスキップされ、もう1つは成功する
    expect(result).toHaveLength(1);
  });

  it("ネットワークエラー時は空配列を返す", async () => {
    const mockFetch = vi.fn(() => {
      return Promise.reject(new Error("Network error"));
    }) as unknown as typeof globalThis.fetch;
    const result = await fetchGoogleFonts(sampleUsedFonts, { fetch: mockFetch });
    expect(result).toEqual([]);
  });
});
