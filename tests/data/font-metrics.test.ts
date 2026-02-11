import { describe, it, expect } from "vitest";
import { getFontMetrics } from "../../src/data/font-metrics.js";

describe("getFontMetrics", () => {
  it("Calibri に対して Carlito ベースのメトリクスを返す", () => {
    const metrics = getFontMetrics("Calibri");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(2048);
    expect(metrics!.widths["A"]).toBe(1185);
  });

  it("Arial に対して Liberation Sans ベースのメトリクスを返す", () => {
    const metrics = getFontMetrics("Arial");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(2048);
    expect(metrics!.widths["A"]).toBe(1366);
  });

  it("Helvetica に対して Liberation Sans ベースのメトリクスを返す", () => {
    const metrics = getFontMetrics("Helvetica");
    expect(metrics).not.toBeNull();
    expect(metrics!.widths["A"]).toBe(1366);
  });

  it("Times New Roman に対して Liberation Serif ベースのメトリクスを返す", () => {
    const metrics = getFontMetrics("Times New Roman");
    expect(metrics).not.toBeNull();
    expect(metrics!.widths["A"]).toBe(1479);
  });

  it("日本語フォント (Meiryo) に対して Noto Sans JP ベースのメトリクスを返す", () => {
    const metrics = getFontMetrics("Meiryo");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(1000);
    expect(metrics!.cjkWidth).toBe(1000);
  });

  it("Yu Gothic に対して Noto Sans JP ベースのメトリクスを返す", () => {
    const metrics = getFontMetrics("Yu Gothic");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(1000);
  });

  it("大文字小文字を無視してマッチングする", () => {
    const metrics1 = getFontMetrics("Calibri");
    const metrics2 = getFontMetrics("calibri");
    expect(metrics1).not.toBeNull();
    expect(metrics2).not.toBeNull();
    expect(metrics1!.widths["A"]).toBe(metrics2!.widths["A"]);
  });

  it("未知のフォントに対して null を返す", () => {
    expect(getFontMetrics("NonExistentFont")).toBeNull();
  });

  it("null に対して null を返す", () => {
    expect(getFontMetrics(null)).toBeNull();
  });

  it("undefined に対して null を返す", () => {
    expect(getFontMetrics(undefined)).toBeNull();
  });

  it("各フォントに ascender と descender が定義されている", () => {
    const calibri = getFontMetrics("Calibri")!;
    expect(calibri.ascender).toBeGreaterThan(0);
    expect(calibri.descender).toBeLessThan(0);

    const arial = getFontMetrics("Arial")!;
    expect(arial.ascender).toBeGreaterThan(0);
    expect(arial.descender).toBeLessThan(0);
  });
});
