import { describe, it, expect } from "vitest";
import { parseEffectList } from "./effect-parser.js";
import { ColorResolver } from "../color/color-resolver.js";
import type { ColorScheme, ColorMap } from "../model/theme.js";

const testScheme: ColorScheme = {
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

const testColorMap: ColorMap = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const resolver = new ColorResolver(testScheme, testColorMap);

describe("parseEffectList", () => {
  it("returns null for null input", () => {
    expect(parseEffectList(null, resolver)).toBeNull();
  });

  it("returns null for undefined input", () => {
    expect(parseEffectList(undefined, resolver)).toBeNull();
  });

  it("returns null for empty object", () => {
    expect(parseEffectList({}, resolver)).toBeNull();
  });

  // --- outerShadow ---

  it("parses outerShdw with all attributes", () => {
    const node = {
      outerShdw: {
        "@_blurRad": "50800",
        "@_dist": "38100",
        "@_dir": "2700000",
        "@_algn": "br",
        "@_rotWithShape": "1",
        srgbClr: { "@_val": "000000", alpha: { "@_val": "50000" } },
      },
    };
    const result = parseEffectList(node, resolver);

    expect(result).not.toBeNull();
    expect(result!.outerShadow).toEqual({
      blurRadius: 50800,
      distance: 38100,
      direction: 45, // 2700000 / 60000
      color: { hex: "#000000", alpha: 0.5 },
      alignment: "br",
      rotateWithShape: true,
    });
    expect(result!.innerShadow).toBeNull();
    expect(result!.glow).toBeNull();
    expect(result!.softEdge).toBeNull();
  });

  it("parses outerShdw with default values", () => {
    const node = {
      outerShdw: {
        srgbClr: { "@_val": "FF0000" },
      },
    };
    const result = parseEffectList(node, resolver);

    expect(result!.outerShadow).toEqual({
      blurRadius: 0,
      distance: 0,
      direction: 0,
      color: { hex: "#FF0000", alpha: 1 },
      alignment: "b",
      rotateWithShape: true,
    });
  });

  it("parses outerShdw rotWithShape='0' as false", () => {
    const node = {
      outerShdw: {
        "@_rotWithShape": "0",
        srgbClr: { "@_val": "FF0000" },
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result!.outerShadow!.rotateWithShape).toBe(false);
  });

  it("returns null when outerShdw has no resolvable color", () => {
    const node = {
      outerShdw: {
        "@_blurRad": "50800",
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result).toBeNull();
  });

  // --- innerShadow ---

  it("parses innerShdw with all attributes", () => {
    const node = {
      innerShdw: {
        "@_blurRad": "63500",
        "@_dist": "38100",
        "@_dir": "5400000",
        srgbClr: { "@_val": "FF0000" },
      },
    };
    const result = parseEffectList(node, resolver);

    expect(result!.innerShadow).toEqual({
      blurRadius: 63500,
      distance: 38100,
      direction: 90, // 5400000 / 60000
      color: { hex: "#FF0000", alpha: 1 },
    });
  });

  it("parses innerShdw with default values", () => {
    const node = {
      innerShdw: {
        srgbClr: { "@_val": "0000FF" },
      },
    };
    const result = parseEffectList(node, resolver);

    expect(result!.innerShadow!.blurRadius).toBe(0);
    expect(result!.innerShadow!.distance).toBe(0);
    expect(result!.innerShadow!.direction).toBe(0);
  });

  it("parses innerShdw with schemeClr", () => {
    const node = {
      innerShdw: {
        "@_blurRad": "63500",
        "@_dist": "38100",
        "@_dir": "5400000",
        schemeClr: { "@_val": "accent1" },
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result!.innerShadow!.color).toEqual({ hex: "#4472C4", alpha: 1 });
  });

  // --- glow ---

  it("parses glow with radius and color", () => {
    const node = {
      glow: {
        "@_rad": "127000",
        srgbClr: { "@_val": "00FF00" },
      },
    };
    const result = parseEffectList(node, resolver);

    expect(result!.glow).toEqual({
      radius: 127000,
      color: { hex: "#00FF00", alpha: 1 },
    });
  });

  it("parses glow with default radius", () => {
    const node = {
      glow: {
        srgbClr: { "@_val": "00FF00" },
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result!.glow!.radius).toBe(0);
  });

  it("returns null when glow has no resolvable color", () => {
    const node = {
      glow: {
        "@_rad": "127000",
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result).toBeNull();
  });

  // --- softEdge ---

  it("parses softEdge with radius", () => {
    const node = {
      softEdge: {
        "@_rad": "63500",
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result!.softEdge).toEqual({ radius: 63500 });
  });

  it("parses softEdge with default radius", () => {
    const node = {
      softEdge: {},
    };
    const result = parseEffectList(node, resolver);
    expect(result!.softEdge!.radius).toBe(0);
  });

  // --- 複数エフェクト ---

  it("parses multiple effects", () => {
    const node = {
      outerShdw: {
        "@_blurRad": "50800",
        srgbClr: { "@_val": "000000" },
      },
      glow: {
        "@_rad": "127000",
        srgbClr: { "@_val": "FF0000" },
      },
      softEdge: {
        "@_rad": "63500",
      },
    };
    const result = parseEffectList(node, resolver);

    expect(result!.outerShadow).not.toBeNull();
    expect(result!.glow).not.toBeNull();
    expect(result!.softEdge).not.toBeNull();
    expect(result!.innerShadow).toBeNull();
  });

  it("parses all four effects", () => {
    const node = {
      outerShdw: { srgbClr: { "@_val": "000000" } },
      innerShdw: { srgbClr: { "@_val": "FF0000" } },
      glow: { srgbClr: { "@_val": "00FF00" } },
      softEdge: { "@_rad": "12700" },
    };
    const result = parseEffectList(node, resolver);

    expect(result!.outerShadow).not.toBeNull();
    expect(result!.innerShadow).not.toBeNull();
    expect(result!.glow).not.toBeNull();
    expect(result!.softEdge).not.toBeNull();
  });

  // --- direction 変換 ---

  it("converts direction 21600000 to 360 degrees", () => {
    const node = {
      outerShdw: {
        "@_dir": "21600000",
        srgbClr: { "@_val": "000000" },
      },
    };
    const result = parseEffectList(node, resolver);
    expect(result!.outerShadow!.direction).toBe(360);
  });
});
