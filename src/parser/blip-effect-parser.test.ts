import { describe, it, expect, vi } from "vitest";
import { parseBlipEffects } from "./blip-effect-parser.js";
import { ColorResolver } from "../color/color-resolver.js";
import type { ColorScheme, ColorMap } from "../model/theme.js";
import { initWarningLogger } from "../warning-logger.js";

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

describe("parseBlipEffects", () => {
  it("returns null for null input", () => {
    expect(parseBlipEffects(null, resolver)).toBeNull();
  });

  it("returns null for empty node", () => {
    expect(parseBlipEffects({}, resolver)).toBeNull();
  });

  it("parses grayscl", () => {
    const result = parseBlipEffects({ grayscl: {} }, resolver);
    expect(result).not.toBeNull();
    expect(result!.grayscale).toBe(true);
  });

  it("parses biLevel with threshold", () => {
    const result = parseBlipEffects({ biLevel: { "@_thresh": "75000" } }, resolver);
    expect(result).not.toBeNull();
    expect(result!.biLevel).toEqual({ threshold: 0.75 });
  });

  it("parses biLevel with default threshold", () => {
    const result = parseBlipEffects({ biLevel: {} }, resolver);
    expect(result!.biLevel!.threshold).toBe(0.5);
  });

  it("parses blur", () => {
    const result = parseBlipEffects({ blur: { "@_rad": "50800", "@_grow": "0" } }, resolver);
    expect(result).not.toBeNull();
    expect(result!.blur).toEqual({ radius: 50800, grow: false });
  });

  it("parses blur with default grow", () => {
    const result = parseBlipEffects({ blur: { "@_rad": "25400" } }, resolver);
    expect(result!.blur!.grow).toBe(true);
  });

  it("parses lum (brightness/contrast)", () => {
    const result = parseBlipEffects(
      { lum: { "@_bright": "70000", "@_contrast": "-30000" } },
      resolver,
    );
    expect(result).not.toBeNull();
    expect(result!.lum).toEqual({ brightness: 0.7, contrast: -0.3 });
  });

  it("parses lum with defaults", () => {
    const result = parseBlipEffects({ lum: {} }, resolver);
    expect(result!.lum).toEqual({ brightness: 0, contrast: 0 });
  });

  it("parses duotone with srgbClr", () => {
    const result = parseBlipEffects(
      {
        duotone: {
          srgbClr: [{ "@_val": "000000" }, { "@_val": "D9C3A5" }],
        },
      },
      resolver,
    );
    expect(result).not.toBeNull();
    expect(result!.duotone!.color1).toEqual({ hex: "#000000", alpha: 1 });
    expect(result!.duotone!.color2).toEqual({ hex: "#D9C3A5", alpha: 1 });
  });

  it("parses duotone with prstClr", () => {
    const result = parseBlipEffects(
      {
        duotone: {
          prstClr: { "@_val": "black" },
          srgbClr: { "@_val": "FF6600" },
        },
      },
      resolver,
    );
    expect(result).not.toBeNull();
    expect(result!.duotone!.color1).toEqual({ hex: "#000000", alpha: 1 });
    expect(result!.duotone!.color2).toEqual({ hex: "#FF6600", alpha: 1 });
  });

  it("returns null for duotone with fewer than 2 colors", () => {
    const result = parseBlipEffects({ duotone: { srgbClr: { "@_val": "000000" } } }, resolver);
    expect(result).toBeNull();
  });

  it("warns for clrChange and returns null if no other effects", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const result = parseBlipEffects({ clrChange: {} }, resolver);
    expect(result).toBeNull();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining("color change effect not implemented"),
    );
    warnSpy.mockRestore();
  });

  it("parses multiple effects", () => {
    const result = parseBlipEffects(
      {
        grayscl: {},
        blur: { "@_rad": "25400" },
      },
      resolver,
    );
    expect(result).not.toBeNull();
    expect(result!.grayscale).toBe(true);
    expect(result!.blur).not.toBeNull();
  });
});
