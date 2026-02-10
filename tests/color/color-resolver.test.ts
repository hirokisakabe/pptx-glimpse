import { describe, it, expect } from "vitest";
import { ColorResolver } from "../../src/color/color-resolver.js";
import type { ColorScheme, ColorMap } from "../../src/model/theme.js";

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

describe("ColorResolver", () => {
  const resolver = new ColorResolver(testScheme, testColorMap);

  it("resolves srgbClr", () => {
    const result = resolver.resolve({ srgbClr: { "@_val": "FF0000" } });
    expect(result).toEqual({ hex: "#FF0000", alpha: 1 });
  });

  it("resolves schemeClr accent1", () => {
    const result = resolver.resolve({ schemeClr: { "@_val": "accent1" } });
    expect(result).toEqual({ hex: "#4472C4", alpha: 1 });
  });

  it("resolves schemeClr via colorMap (tx1 → dk1)", () => {
    const result = resolver.resolve({ schemeClr: { "@_val": "tx1" } });
    expect(result).toEqual({ hex: "#000000", alpha: 1 });
  });

  it("resolves schemeClr via colorMap (bg1 → lt1)", () => {
    const result = resolver.resolve({ schemeClr: { "@_val": "bg1" } });
    expect(result).toEqual({ hex: "#FFFFFF", alpha: 1 });
  });

  it("resolves sysClr with lastClr", () => {
    const result = resolver.resolve({
      sysClr: { "@_val": "windowText", "@_lastClr": "000000" },
    });
    expect(result).toEqual({ hex: "#000000", alpha: 1 });
  });

  it("applies alpha", () => {
    const result = resolver.resolve({
      srgbClr: { "@_val": "FF0000", alpha: { "@_val": "50000" } },
    });
    expect(result?.hex).toBe("#FF0000");
    expect(result?.alpha).toBe(0.5);
  });

  it("returns null for empty node", () => {
    expect(resolver.resolve(null)).toBeNull();
    expect(resolver.resolve({})).toBeNull();
  });
});
