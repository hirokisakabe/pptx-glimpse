import { describe, it, expect, vi } from "vitest";
import { CanvasTextMeasurer } from "./canvas-text-measurer.js";

function createMockCanvas() {
  const measureText = vi.fn().mockReturnValue({
    width: 100,
    fontBoundingBoxAscent: 16,
    fontBoundingBoxDescent: 4,
  });
  const ctx = { font: "", measureText };
  const canvas = {
    getContext: (_id: "2d") => ctx,
  };
  return { canvas, ctx, measureText };
}

describe("CanvasTextMeasurer", () => {
  it("measureTextWidth は Canvas measureText を呼ぶ", () => {
    const { canvas, measureText } = createMockCanvas();
    const measurer = new CanvasTextMeasurer(canvas);
    const width = measurer.measureTextWidth("Hello", 18, false, "Arial");
    expect(measureText).toHaveBeenCalledWith("Hello");
    expect(width).toBe(100);
  });

  it("太字の場合フォントに bold が設定される", () => {
    const { canvas, ctx } = createMockCanvas();
    const measurer = new CanvasTextMeasurer(canvas);
    measurer.measureTextWidth("Hello", 18, true, "Arial");
    expect(ctx.font).toContain("bold");
  });

  it("getLineHeightRatio は fontBoundingBox から計算する", () => {
    const { canvas } = createMockCanvas();
    const measurer = new CanvasTextMeasurer(canvas);
    const ratio = measurer.getLineHeightRatio("Arial");
    expect(ratio).toBeCloseTo((16 + 4) / 16, 5);
  });

  it("fontBoundingBox がない場合は 1.2 にフォールバックする", () => {
    const canvas = {
      getContext: (_id: "2d") => ({
        font: "",
        measureText: (_text: string) => ({ width: 100 }),
      }),
    };
    const measurer = new CanvasTextMeasurer(canvas);
    expect(measurer.getLineHeightRatio("Arial")).toBe(1.2);
  });

  it("getContext が null を返す場合はエラーを投げる", () => {
    const canvas = {
      getContext: (_id: "2d") => null,
    };
    expect(() => new CanvasTextMeasurer(canvas)).toThrow("Failed to get 2D context");
  });
});
