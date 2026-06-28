import { describe, expect, it } from "vitest";

import { getFontMetrics, getMetricsFallbackFont } from "./font-metrics.js";

describe("getFontMetrics", () => {
  it("covers font-metrics behavior 1", () => {
    const metrics = getFontMetrics("Calibri");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(2048);
    expect(metrics!.widths["A"]).toBe(1185);
  });

  it("covers font-metrics behavior 2", () => {
    const metrics = getFontMetrics("Arial");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(2048);
    expect(metrics!.widths["A"]).toBe(1366);
  });

  it("covers font-metrics behavior 3", () => {
    const metrics = getFontMetrics("Helvetica");
    expect(metrics).not.toBeNull();
    expect(metrics!.widths["A"]).toBe(1366);
  });

  it("covers font-metrics behavior 4", () => {
    const metrics = getFontMetrics("Times New Roman");
    expect(metrics).not.toBeNull();
    expect(metrics!.widths["A"]).toBe(1479);
  });

  it("covers font-metrics behavior 5", () => {
    const metrics = getFontMetrics("Meiryo");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(1000);
    expect(metrics!.cjkWidth).toBe(1000);
  });

  it("covers font-metrics behavior 6", () => {
    const metrics = getFontMetrics("Yu Gothic");
    expect(metrics).not.toBeNull();
    expect(metrics!.unitsPerEm).toBe(1000);
  });

  it("covers font-metrics behavior 7", () => {
    const metrics1 = getFontMetrics("Calibri");
    const metrics2 = getFontMetrics("calibri");
    expect(metrics1).not.toBeNull();
    expect(metrics2).not.toBeNull();
    expect(metrics1!.widths["A"]).toBe(metrics2!.widths["A"]);
  });

  it("covers font-metrics behavior 8", () => {
    expect(getFontMetrics("NonExistentFont")).toBeNull();
  });

  it("covers font-metrics behavior 9", () => {
    expect(getFontMetrics(null)).toBeNull();
  });

  it("covers font-metrics behavior 10", () => {
    expect(getFontMetrics(undefined)).toBeNull();
  });

  it("covers font-metrics behavior 11", () => {
    const calibri = getFontMetrics("Calibri")!;
    expect(calibri.ascender).toBeGreaterThan(0);
    expect(calibri.descender).toBeLessThan(0);

    const arial = getFontMetrics("Arial")!;
    expect(arial.ascender).toBeGreaterThan(0);
    expect(arial.descender).toBeLessThan(0);
  });
});

describe("getMetricsFallbackFont", () => {
  it("covers font-metrics behavior 12", () => {
    expect(getMetricsFallbackFont("Calibri")).toBe("Carlito");
  });

  it("covers font-metrics behavior 13", () => {
    expect(getMetricsFallbackFont("Arial")).toBe("Liberation Sans");
  });

  it("covers font-metrics behavior 14", () => {
    expect(getMetricsFallbackFont("Helvetica")).toBe("Liberation Sans");
  });

  it("covers font-metrics behavior 15", () => {
    expect(getMetricsFallbackFont("Times New Roman")).toBe("Liberation Serif");
  });

  it("covers font-metrics behavior 16", () => {
    expect(getMetricsFallbackFont("Meiryo")).toBe("Noto Sans JP");
  });

  it("covers font-metrics behavior 17", () => {
    expect(getMetricsFallbackFont("Yu Gothic")).toBe("Noto Sans JP");
  });

  it("covers font-metrics behavior 18", () => {
    expect(getMetricsFallbackFont("Noto Sans JP")).toBe("Noto Sans JP");
  });

  it("covers font-metrics behavior 19", () => {
    expect(getMetricsFallbackFont("calibri")).toBe("Carlito");
  });

  it("covers font-metrics behavior 20", () => {
    expect(getMetricsFallbackFont("NonExistentFont")).toBeNull();
  });

  it("covers font-metrics behavior 21", () => {
    expect(getMetricsFallbackFont(null)).toBeNull();
  });

  it("covers font-metrics behavior 22", () => {
    expect(getMetricsFallbackFont(undefined)).toBeNull();
  });
});
