import type { TextMeasurer } from "./text-measurer.js";

const PX_PER_PT = 96 / 72;

interface TextMetricsLike {
  width: number;
  fontBoundingBoxAscent?: number;
  fontBoundingBoxDescent?: number;
}

interface CanvasContext {
  font: string;
  measureText(text: string): TextMetricsLike;
}

interface CanvasLike {
  getContext(contextId: "2d"): CanvasContext | null;
}

export class CanvasTextMeasurer implements TextMeasurer {
  private ctx: CanvasContext;

  /**
   * @param canvas - Canvas 要素または getContext("2d") をサポートするオブジェクト。
   *   省略した場合は document.createElement("canvas") で生成する（ブラウザ環境のみ）。
   */
  constructor(canvas?: CanvasLike) {
    const cvs =
      canvas ??
      (
        (globalThis as Record<string, unknown>).document as {
          createElement(tag: string): CanvasLike;
        }
      ).createElement("canvas");
    const ctx = cvs.getContext("2d");
    if (!ctx) throw new Error("Failed to get 2D context");
    this.ctx = ctx;
  }

  measureTextWidth(
    text: string,
    fontSizePt: number,
    bold: boolean,
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
  ): number {
    const fontSizePx = fontSizePt * PX_PER_PT;
    const weight = bold ? "bold" : "normal";
    const families = [fontFamily, fontFamilyEa, "sans-serif"].filter(Boolean).join(", ");
    this.ctx.font = `${weight} ${fontSizePx}px ${families}`;
    return this.ctx.measureText(text).width;
  }

  getLineHeightRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
    const families = [fontFamily, fontFamilyEa, "sans-serif"].filter(Boolean).join(", ");
    this.ctx.font = `16px ${families}`;
    const metrics = this.ctx.measureText("Hg漢");
    if (
      metrics.fontBoundingBoxAscent !== undefined &&
      metrics.fontBoundingBoxDescent !== undefined
    ) {
      return (metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent) / 16;
    }
    return 1.2;
  }
}
