import { describe, it, expect } from "vitest";
import { resolveShapeStyle } from "./style-reference-resolver.js";
import { ColorResolver } from "../color/color-resolver.js";
import type { FormatScheme } from "../model/theme.js";
import type { XmlNode } from "./xml-parser.js";

const colorScheme = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#44546A",
  lt2: "#E7E6E6",
  accent1: "#4472C4",
  accent2: "#ED7D31",
  accent3: "#A5A5A5",
  accent4: "#FFC000",
  accent5: "#5B9BD5",
  accent6: "#70AD47",
  hlink: "#0563C1",
  folHlink: "#954F72",
};

const colorMap = {
  bg1: "lt1" as const,
  tx1: "dk1" as const,
  bg2: "lt2" as const,
  tx2: "dk2" as const,
  accent1: "accent1" as const,
  accent2: "accent2" as const,
  accent3: "accent3" as const,
  accent4: "accent4" as const,
  accent5: "accent5" as const,
  accent6: "accent6" as const,
  hlink: "hlink" as const,
  folHlink: "folHlink" as const,
};

const colorResolver = new ColorResolver(colorScheme, colorMap);

const fmtScheme: FormatScheme = {
  fillStyles: [
    { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
    {
      type: "gradient",
      stops: [
        { position: 0, color: { hex: "#00FF00", alpha: 1 } },
        { position: 1, color: { hex: "#0000FF", alpha: 1 } },
      ],
      angle: 90,
      gradientType: "linear",
    },
  ],
  lnStyles: [
    {
      width: 6350,
      fill: { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
      dashStyle: "solid",
      lineCap: undefined,
      lineJoin: undefined,
      headEnd: null,
      tailEnd: null,
    },
    {
      width: 12700,
      fill: { type: "solid", color: { hex: "#00FF00", alpha: 1 } },
      dashStyle: "solid",
      lineCap: undefined,
      lineJoin: undefined,
      headEnd: null,
      tailEnd: null,
    },
  ],
  effectStyles: [
    null,
    {
      outerShadow: {
        blurRadius: 57150,
        distance: 19050,
        direction: 90,
        color: { hex: "#000000", alpha: 0.63 },
        alignment: "ctr",
        rotateWithShape: false,
      },
      innerShadow: null,
      glow: null,
      softEdge: null,
    },
  ],
  bgFillStyles: [{ type: "solid", color: { hex: "#AABBCC", alpha: 1 } }],
};

describe("resolveShapeStyle", () => {
  it("returns null when styleNode is undefined", () => {
    expect(resolveShapeStyle(undefined, fmtScheme, colorResolver)).toBeNull();
  });

  it("returns null when fmtScheme is undefined", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "1", schemeClr: { "@_val": "accent1" } },
    };
    expect(resolveShapeStyle(styleNode, undefined, colorResolver)).toBeNull();
  });

  it("resolves fillRef idx=1 to first fill style", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "1", schemeClr: { "@_val": "accent1" } },
      lnRef: { "@_idx": "0" },
      effectRef: { "@_idx": "0" },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.fill).not.toBeNull();
    // fillRef has a color override via schemeClr, so the solid fill color is overridden
    expect(result!.fill!.type).toBe("solid");
    if (result!.fill!.type === "solid") {
      expect(result!.fill!.color.hex).toBe("#4472C4"); // accent1
    }
  });

  it("resolves fillRef idx=0 as no fill", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "0" },
      lnRef: { "@_idx": "0" },
      effectRef: { "@_idx": "0" },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.fill).toBeNull();
    expect(result!.outline).toBeNull();
  });

  it("resolves lnRef idx=2 to second line style with color override", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "0" },
      lnRef: { "@_idx": "2", srgbClr: { "@_val": "ABCDEF" } },
      effectRef: { "@_idx": "0" },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.outline).not.toBeNull();
    expect(result!.outline!.width).toBe(12700);
    expect(result!.outline!.fill!.type).toBe("solid");
    if (result!.outline!.fill!.type === "solid") {
      expect(result!.outline!.fill!.color.hex).toBe("#ABCDEF");
    }
  });

  it("resolves effectRef idx=1 to second effect style", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "0" },
      lnRef: { "@_idx": "0" },
      effectRef: { "@_idx": "1" },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.effects).not.toBeNull();
    expect(result!.effects!.outerShadow).toBeDefined();
  });

  it("resolves fillRef idx >= 1000 from bgFillStyleLst", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "1001", schemeClr: { "@_val": "accent2" } },
      lnRef: { "@_idx": "0" },
      effectRef: { "@_idx": "0" },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.fill).not.toBeNull();
    expect(result!.fill!.type).toBe("solid");
    if (result!.fill!.type === "solid") {
      expect(result!.fill!.color.hex).toBe("#ED7D31"); // accent2
    }
  });

  it("resolves fontRef", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "0" },
      lnRef: { "@_idx": "0" },
      effectRef: { "@_idx": "0" },
      fontRef: { "@_idx": "minor", schemeClr: { "@_val": "dk1" } },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.fontRef).toBeDefined();
    expect(result!.fontRef!.idx).toBe("minor");
    expect(result!.fontRef!.color).not.toBeNull();
    expect(result!.fontRef!.color!.hex).toBe("#000000"); // dk1
  });

  it("overrides gradient fill colors with ref color", () => {
    const styleNode: XmlNode = {
      fillRef: { "@_idx": "2", srgbClr: { "@_val": "123456" } },
      lnRef: { "@_idx": "0" },
      effectRef: { "@_idx": "0" },
    };
    const result = resolveShapeStyle(styleNode, fmtScheme, colorResolver);

    expect(result).not.toBeNull();
    expect(result!.fill).not.toBeNull();
    expect(result!.fill!.type).toBe("gradient");
    if (result!.fill!.type === "gradient") {
      // All stops should have the override color
      for (const stop of result!.fill!.stops) {
        expect(stop.color.hex).toBe("#123456");
      }
    }
  });
});
