import { describe, it, expect } from "vitest";
import { renderFillAttrs, renderOutlineAttrs } from "./fill-renderer.js";

describe("renderFillAttrs", () => {
  it("renders null fill as none", () => {
    const result = renderFillAttrs(null);
    expect(result.attrs).toBe('fill="none"');
    expect(result.defs).toBe("");
  });

  it("renders noFill", () => {
    const result = renderFillAttrs({ type: "none" });
    expect(result.attrs).toBe('fill="none"');
  });

  it("renders solid fill", () => {
    const result = renderFillAttrs({
      type: "solid",
      color: { hex: "#FF0000", alpha: 1 },
    });
    expect(result.attrs).toBe('fill="#FF0000"');
    expect(result.defs).toBe("");
  });

  it("renders solid fill with alpha", () => {
    const result = renderFillAttrs({
      type: "solid",
      color: { hex: "#FF0000", alpha: 0.5 },
    });
    expect(result.attrs).toContain('fill="#FF0000"');
    expect(result.attrs).toContain('fill-opacity="0.5"');
  });

  it("renders gradient fill with defs", () => {
    const result = renderFillAttrs({
      type: "gradient",
      angle: 90,
      gradientType: "linear",
      stops: [
        { position: 0, color: { hex: "#FF0000", alpha: 1 } },
        { position: 1, color: { hex: "#0000FF", alpha: 1 } },
      ],
    });
    expect(result.attrs).toContain("url(#grad-");
    expect(result.defs).toContain("<linearGradient");
    expect(result.defs).toContain("#FF0000");
    expect(result.defs).toContain("#0000FF");
  });

  it("renders radial gradient fill", () => {
    const result = renderFillAttrs({
      type: "gradient",
      angle: 0,
      gradientType: "radial",
      centerX: 0.5,
      centerY: 0.5,
      stops: [
        { position: 0, color: { hex: "#FF0000", alpha: 1 } },
        { position: 1, color: { hex: "#0000FF", alpha: 1 } },
      ],
    });
    expect(result.attrs).toContain("url(#grad-");
    expect(result.defs).toContain("<radialGradient");
    expect(result.defs).toContain('cx="50%"');
    expect(result.defs).toContain('cy="50%"');
    expect(result.defs).toContain("#FF0000");
    expect(result.defs).toContain("#0000FF");
  });

  it("renders image fill with pattern", () => {
    const result = renderFillAttrs({
      type: "image",
      imageData: "dGVzdA==",
      mimeType: "image/png",
      tile: null,
    });
    expect(result.attrs).toContain("url(#imgfill-");
    expect(result.defs).toContain("<pattern");
    expect(result.defs).toContain('patternContentUnits="objectBoundingBox"');
    expect(result.defs).toContain("<image");
    expect(result.defs).toContain("data:image/png;base64,dGVzdA==");
  });

  it("renders EMF fill as gray placeholder", () => {
    const result = renderFillAttrs({
      type: "image",
      imageData: "dGVzdA==",
      mimeType: "image/emf",
      tile: null,
    });
    expect(result.attrs).toBe('fill="#E0E0E0"');
    expect(result.defs).toBe("");
  });

  it("renders image fill with tile", () => {
    const result = renderFillAttrs({
      type: "image",
      imageData: "dGVzdA==",
      mimeType: "image/png",
      tile: { tx: 0, ty: 0, sx: 0.5, sy: 0.5, flip: "none", align: "tl" },
    });
    expect(result.attrs).toContain("url(#imgfill-");
    expect(result.defs).toContain("patternUnits=");
    expect(result.defs).toContain('width="50%"');
    expect(result.defs).toContain('height="50%"');
  });

  it("renders pattern fill", () => {
    const result = renderFillAttrs({
      type: "pattern",
      preset: "ltDnDiag",
      foregroundColor: { hex: "#4472C4", alpha: 1 },
      backgroundColor: { hex: "#FFFFFF", alpha: 1 },
    });
    expect(result.attrs).toContain("url(#patt-");
    expect(result.defs).toContain("<pattern");
    expect(result.defs).toContain('patternUnits="userSpaceOnUse"');
    expect(result.defs).toContain("#FFFFFF");
    expect(result.defs).toContain("#4472C4");
  });

  it("renders unknown pattern preset as solid foreground", () => {
    const result = renderFillAttrs({
      type: "pattern",
      preset: "unknownPattern",
      foregroundColor: { hex: "#4472C4", alpha: 1 },
      backgroundColor: { hex: "#FFFFFF", alpha: 1 },
    });
    expect(result.attrs).toBe('fill="#4472C4"');
    expect(result.defs).toBe("");
  });
});

describe("renderOutlineAttrs", () => {
  it("renders null outline as stroke none", () => {
    const result = renderOutlineAttrs(null);
    expect(result.attrs).toBe('stroke="none"');
    expect(result.defs).toBe("");
  });

  it("renders outline with color", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "solid",
      headEnd: null,
      tailEnd: null,
    });
    expect(result.attrs).toContain('stroke="#000000"');
    expect(result.attrs).toContain("stroke-width=");
    expect(result.defs).toBe("");
  });

  it("renders dash style", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "dash",
      headEnd: null,
      tailEnd: null,
    });
    expect(result.attrs).toContain("stroke-dasharray=");
  });

  it("renders gradient outline with defs", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: {
        type: "gradient",
        stops: [
          { position: 0, color: { hex: "#FF0000", alpha: 1 } },
          { position: 1, color: { hex: "#0000FF", alpha: 1 } },
        ],
        angle: 90,
        gradientType: "linear",
      },
      dashStyle: "solid",
      headEnd: null,
      tailEnd: null,
    });
    expect(result.attrs).toContain('stroke="url(#grad-');
    expect(result.defs).toContain("<linearGradient");
    expect(result.defs).toContain("#FF0000");
    expect(result.defs).toContain("#0000FF");
  });

  it("renders custom dash pattern", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "solid",
      customDash: [3, 1, 1, 1],
      headEnd: null,
      tailEnd: null,
    });
    expect(result.attrs).toContain("stroke-dasharray=");
  });

  it("custom dash overrides prstDash", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "dash",
      customDash: [3, 1],
      headEnd: null,
      tailEnd: null,
    });
    // customDash が適用される (prstDash ではなく)
    const widthPx = (12700 / 914400) * 96;
    expect(result.attrs).toContain(`stroke-dasharray="${3 * widthPx} ${1 * widthPx}"`);
  });
});
