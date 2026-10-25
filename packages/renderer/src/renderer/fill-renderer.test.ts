import { describe, expect, it } from "vitest";

import { createRepresentativeWmf } from "../../../../vrt/snapshot/fixtures-src/images.js";
import type { Outline } from "../model/line.js";
import { createWarningLogger } from "../warning-logger.js";
import { renderFillAttrs, renderMarkers, renderOutlineAttrs } from "./fill-renderer.js";
import { createRendererContext } from "./render-context.js";

describe("renderFillAttrs", () => {
  it("converts a WMF image fill to an SVG pattern", () => {
    const warnings = createWarningLogger("warn");
    const result = renderFillAttrs(
      {
        type: "image",
        mimeType: "image/wmf",
        imageData: createRepresentativeWmf().toString("base64"),
        tile: null,
      },
      createRendererContext({ warningLogger: warnings }),
    );

    expect(result.attrs).toContain('fill="url(#imgfill-');
    expect(result.defs).toContain('viewBox="0 0 1000 1000"');
    expect(result.defs).toContain(">WMF</text>");
    expect(warnings.getWarningEntries()).toEqual([]);
  });
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

  it("renders large dash-dot-dot style", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "lgDashDotDot",
      headEnd: null,
      tailEnd: null,
    });
    expect(result.attrs).toContain('stroke-dasharray="10.666666666666666 4 1.3333333333333333 4');
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
    // customDash is applied (instead of prstDash)
    const widthPx = (12700 / 914400) * 96;
    expect(result.attrs).toContain(`stroke-dasharray="${3 * widthPx} ${1 * widthPx}"`);
  });
});

describe("renderMarkers", () => {
  const solidBlack = { type: "solid", color: { hex: "#000000", alpha: 1 } } as const;
  const triangle = { type: "triangle", width: "med", length: "med" } as const;

  const outlineWith = (ends: Partial<Pick<Outline, "headEnd" | "tailEnd">>): Outline => ({
    width: 12700,
    fill: solidBlack,
    dashStyle: "solid",
    headEnd: null,
    tailEnd: null,
    ...ends,
  });

  it("renders no markers when both endpoints are absent", () => {
    const result = renderMarkers(outlineWith({}));
    expect(result).toEqual({ defs: "", startAttr: "", endAttr: "" });
  });

  it("anchors the end marker on its tip without mirroring", () => {
    const result = renderMarkers(outlineWith({ tailEnd: triangle }));
    // ARROW_SIZE_MAP.med === 8, and every arrow path draws its tip at x=mw.
    expect(result.defs).toContain('refX="8"');
    expect(result.defs).not.toContain("scale(-1,1)");
    expect(result.endAttr).toMatch(/^marker-end="url\(#marker-/);
    expect(result.startAttr).toBe("");
  });

  it("mirrors the start marker so the arrow points away from the line", () => {
    const result = renderMarkers(outlineWith({ headEnd: triangle }));
    // orient="auto" aligns the marker's +x axis with the path direction, which at the
    // start vertex points into the line. Mirroring turns the tip outwards, and refX=0
    // keeps the mirrored tip anchored on the start point.
    expect(result.defs).toContain('refX="0"');
    expect(result.defs).toContain('transform="translate(8,0) scale(-1,1)"');
    expect(result.startAttr).toMatch(/^marker-start="url\(#marker-/);
    expect(result.endAttr).toBe("");
  });

  it("keeps both endpoints independent", () => {
    const result = renderMarkers(outlineWith({ headEnd: triangle, tailEnd: triangle }));
    expect(result.defs.match(/<marker /g)).toHaveLength(2);
    expect(result.defs).toContain('refX="0"');
    expect(result.defs).toContain('refX="8"');
    expect(result.startAttr).not.toBe("");
    expect(result.endAttr).not.toBe("");
  });

  it("does not mirror the symmetric oval endpoint", () => {
    const oval = { type: "oval", width: "med", length: "med" } as const;
    const start = renderMarkers(outlineWith({ headEnd: oval }));
    const end = renderMarkers(outlineWith({ tailEnd: oval }));
    expect(start.defs).not.toContain("scale(-1,1)");
    expect(start.defs).toContain('refX="4"');
    expect(end.defs).toContain('refX="4"');
  });

  it("skips endpoints of type none", () => {
    const none = { type: "none", width: "med", length: "med" } as const;
    const result = renderMarkers(outlineWith({ headEnd: none, tailEnd: none }));
    expect(result).toEqual({ defs: "", startAttr: "", endAttr: "" });
  });
});
