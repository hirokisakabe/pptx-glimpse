import { describe, it, expect } from "vitest";
import { measureTextWidth } from "../../src/utils/text-measure.js";

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
