import { describe, it, expect } from "vitest";
import { measureTextWidth, getLineHeightRatio } from "./text-measure.js";

const PX_PER_PT = 96 / 72;

describe("measureTextWidth", () => {
  it("空文字列の幅は 0", () => {
    expect(measureTextWidth("", 18, false)).toBe(0);
  });

  it("ASCII テキストの幅を推定する", () => {
    const width = measureTextWidth("Hello", 18, false);
    // 'H','e','o' = normal(0.6), 'l','l' = narrow(0.3)
    // (3 * 0.6 + 2 * 0.3) * 18 * (96/72) = (1.8 + 0.6) * 24 = 57.6
    expect(width).toBeCloseTo(57.6, 1);
  });

  it("CJK テキストの幅を推定する", () => {
    const width = measureTextWidth("漢字", 18, false);
    // 2 * 1.0 * 18 * (96/72) = 48
    expect(width).toBeCloseTo(2 * 1.0 * 18 * PX_PER_PT, 1);
  });

  it("混合テキストの幅を推定する", () => {
    const width = measureTextWidth("A漢", 18, false);
    // A=normal(0.6) + 漢=wide(1.0) = 1.6 * 18 * (96/72) = 38.4
    expect(width).toBeCloseTo(1.6 * 18 * PX_PER_PT, 1);
  });

  it("太字の場合は幅が増加する", () => {
    const normalWidth = measureTextWidth("Test", 18, false);
    const boldWidth = measureTextWidth("Test", 18, true);
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("スペースの幅を推定する", () => {
    const width = measureTextWidth(" ", 18, false);
    // narrow(0.3) * 18 * (96/72) = 7.2
    expect(width).toBeCloseTo(0.3 * 18 * PX_PER_PT, 1);
  });

  it("ひらがなを wide として推定する", () => {
    const width = measureTextWidth("あ", 18, false);
    expect(width).toBeCloseTo(1.0 * 18 * PX_PER_PT, 1);
  });

  it("カタカナを wide として推定する", () => {
    const width = measureTextWidth("ア", 18, false);
    expect(width).toBeCloseTo(1.0 * 18 * PX_PER_PT, 1);
  });

  it("フォントサイズに比例する", () => {
    const width12 = measureTextWidth("A", 12, false);
    const width24 = measureTextWidth("A", 24, false);
    expect(width24).toBeCloseTo(width12 * 2, 1);
  });
});

describe("measureTextWidth with font metrics", () => {
  it("Calibri のメトリクスで ASCII テキストの幅を計算する", () => {
    // Carlito: A=1185, unitsPerEm=2048
    // 幅 = 1185 / 2048 * 18 * (96/72)
    const width = measureTextWidth("A", 18, false, "Calibri");
    const expected = (1185 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("ヒューリスティックとは異なる値を返す", () => {
    const metricsWidth = measureTextWidth("A", 18, false, "Calibri");
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).not.toBeCloseTo(heuristicWidth, 0);
  });

  it("未知のフォントではヒューリスティックにフォールバックする", () => {
    const metricsWidth = measureTextWidth("A", 18, false, "UnknownFont");
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).toBeCloseTo(heuristicWidth, 5);
  });

  it("fontFamily が null の場合はヒューリスティックにフォールバックする", () => {
    const metricsWidth = measureTextWidth("A", 18, false, null);
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).toBeCloseTo(heuristicWidth, 5);
  });

  it("Arial のメトリクスで計算する", () => {
    // LiberationSans: A=1366, unitsPerEm=2048
    const width = measureTextWidth("A", 18, false, "Arial");
    const expected = (1366 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("CJK テキストは cjkWidth を使用する", () => {
    // Carlito: cjkWidth=2048, unitsPerEm=2048 → 1.0 * fontSizePx
    const width = measureTextWidth("漢", 18, false, "Calibri");
    const expected = (2048 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("太字はメトリクスベースの幅にも BOLD_FACTOR を適用する", () => {
    const normalWidth = measureTextWidth("Test", 18, false, "Calibri");
    const boldWidth = measureTextWidth("Test", 18, true, "Calibri");
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("複数文字の幅を正しく合算する", () => {
    // Carlito: H=1276, e=1019, l=470, l=470, o=1080
    const width = measureTextWidth("Hello", 18, false, "Calibri");
    const expected = ((1276 + 1019 + 470 + 470 + 1080) / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("メトリクスに無い文字は defaultWidth を使用する", () => {
    // Carlito: defaultWidth=991, unitsPerEm=2048
    // U+0100 (Ā) はメトリクスに含まれない Latin 拡張文字
    const width = measureTextWidth("\u0100", 18, false, "Calibri");
    const expected = (991 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("measureTextWidth with fontFamilyEa", () => {
  it("CJK 文字に ea フォントメトリクスを使用する", () => {
    // NotoSansJP: cjkWidth=1000, unitsPerEm=1000
    const width = measureTextWidth("漢", 18, false, "Calibri", "Noto Sans JP");
    const expected = (1000 / 1000) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("latin 文字には latin フォントメトリクスを使用する", () => {
    // Carlito: A=1185, unitsPerEm=2048
    const width = measureTextWidth("A", 18, false, "Calibri", "Noto Sans JP");
    const expected = (1185 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("混在テキストで文字ごとに異なるメトリクスを使用する", () => {
    // A: Carlito (1185/2048), 漢: NotoSansJP (1000/1000)
    const width = measureTextWidth("A漢", 18, false, "Calibri", "Noto Sans JP");
    const expectedLatin = (1185 / 2048) * 18 * PX_PER_PT;
    const expectedEa = (1000 / 1000) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expectedLatin + expectedEa, 1);
  });

  it("fontFamilyEa が null の場合は latin メトリクスで CJK も計測する", () => {
    // Carlito: cjkWidth=2048, unitsPerEm=2048
    const width = measureTextWidth("漢", 18, false, "Calibri", null);
    const expected = (2048 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("getLineHeightRatio", () => {
  it("Calibri (Carlito) の行高さ比率を返す", () => {
    // Carlito: ascender=1950, descender=-550, unitsPerEm=2048
    // (1950 + 550) / 2048 = 1.220703125
    const ratio = getLineHeightRatio("Calibri");
    expect(ratio).toBeCloseTo((1950 + 550) / 2048, 5);
  });

  it("Arial (LiberationSans) の行高さ比率を返す", () => {
    // LiberationSans: ascender=1854, descender=-434, unitsPerEm=2048
    const ratio = getLineHeightRatio("Arial");
    expect(ratio).toBeCloseTo((1854 + 434) / 2048, 5);
  });

  it("NotoSansJP の行高さ比率を返す", () => {
    // NotoSansJP: ascender=1160, descender=-288, unitsPerEm=1000
    const ratio = getLineHeightRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo((1160 + 288) / 1000, 5);
  });

  it("fontFamily を優先して使用する", () => {
    const ratio = getLineHeightRatio("Calibri", "Meiryo");
    expect(ratio).toBeCloseTo((1950 + 550) / 2048, 5);
  });

  it("fontFamily が null の場合は fontFamilyEa を使用する", () => {
    const ratio = getLineHeightRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo((1160 + 288) / 1000, 5);
  });

  it("未知のフォントの場合はデフォルト値 1.2 を返す", () => {
    const ratio = getLineHeightRatio("UnknownFont");
    expect(ratio).toBe(1.2);
  });

  it("両方 null の場合はデフォルト値 1.2 を返す", () => {
    const ratio = getLineHeightRatio(null, null);
    expect(ratio).toBe(1.2);
  });

  it("引数なしの場合はデフォルト値 1.2 を返す", () => {
    const ratio = getLineHeightRatio();
    expect(ratio).toBe(1.2);
  });
});
