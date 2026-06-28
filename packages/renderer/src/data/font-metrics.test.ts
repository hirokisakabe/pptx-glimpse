import { describe, expect, it } from "vitest";

import { getFontMetrics, getMetricsFallbackFont } from "./font-metrics.js";

describe("getFontMetrics", () => {
  it("Return Carlito-based metrics for Calibri", () => {
    const metrics = getFontMetrics("Calibri");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(2048);
    expect(metrics!.widths["A"]).toBe(1185);
  });

  it("Return Liberation Sans-based metrics for Arial", () => {
    const metrics = getFontMetrics("Arial");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(2048);
    expect(metrics!.widths["A"]).toBe(1366);
  });

  it("Returning Liberation Sans-based metrics for Helvetica", () => {
    const metrics = getFontMetrics("Helvetica");
    expect(metrics).not.toBeNull();
    expect(metrics!.widths["A"]).toBe(1366);
  });

  it("Return Liberation Serif-based metrics for Times New Roman", () => {
    const metrics = getFontMetrics("Times New Roman");
    expect(metrics).not.toBeNull();
    expect(metrics!.widths["A"]).toBe(1479);
  });

  it("Return Noto Sans JP based metrics for Japanese font (Meiryo)", () => {
    const metrics = getFontMetrics("Meiryo");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(1000);
    expect(metrics!.cjkWidth).toBe(1000);
  });

  it("Return Noto Sans JP based metrics for Yu Gothic", () => {
    const metrics = getFontMetrics("Yu Gothic");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(1000);
  });

  it("Match ignoring case", () => {
    const metrics1 = getFontMetrics("Calibri");
    const metrics2 = getFontMetrics("calibri");
    expect(metrics1).not.toBeNull();
    expect(metrics2).not.toBeNull();
    expect(metrics1!.widths["A"]).toBe(metrics2!.widths["A"]);
  });

  it("Return null for unknown fonts", () => {
    expect(getFontMetrics("NonExistentFont")).toBeNull();
  });

  it("return null for null", () => {
    expect(getFontMetrics(null)).toBeNull();
  });

  it("Return null for undefined", () => {
    expect(getFontMetrics(undefined)).toBeNull();
  });

  it("Each font has an ascender and descender defined", () => {
    const calibri = getFontMetrics("Calibri")!;
    expect(calibri.ascender).toBeGreaterThan(0);
    expect(calibri.descender).toBeLessThan(0);

    const arial = getFontMetrics("Arial")!;
    expect(arial.ascender).toBeGreaterThan(0);
    expect(arial.descender).toBeLessThan(0);
  });
});

describe("getMetricsFallbackFont", () => {
  it("Return Carlito for Calibri", () => {
    expect(getMetricsFallbackFont("Calibri")).toBe("Carlito");
  });

  it("Return Liberation Sans for Arial", () => {
    expect(getMetricsFallbackFont("Arial")).toBe("Liberation Sans");
  });

  it("Return Liberation Sans for Helvetica", () => {
    expect(getMetricsFallbackFont("Helvetica")).toBe("Liberation Sans");
  });

  it("Return Liberation Serif for Times New Roman", () => {
    expect(getMetricsFallbackFont("Times New Roman")).toBe("Liberation Serif");
  });

  it("Return Noto Sans JP to Meiryo", () => {
    expect(getMetricsFallbackFont("Meiryo")).toBe("Noto Sans JP");
  });

  it("Return Noto Sans JP for Yu Gothic", () => {
    expect(getMetricsFallbackFont("Yu Gothic")).toBe("Noto Sans JP");
  });

  it("Return Noto Sans JP for Noto Sans JP", () => {
    expect(getMetricsFallbackFont("Noto Sans JP")).toBe("Noto Sans JP");
  });

  it("Match ignoring case", () => {
    expect(getMetricsFallbackFont("calibri")).toBe("Carlito");
  });

  it("Return null for unknown fonts", () => {
    expect(getMetricsFallbackFont("NonExistentFont")).toBeNull();
  });

  it("return null for null", () => {
    expect(getMetricsFallbackFont(null)).toBeNull();
  });

  it("Return null for undefined", () => {
    expect(getMetricsFallbackFont(undefined)).toBeNull();
  });
});
