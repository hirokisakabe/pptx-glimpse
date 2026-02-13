import { describe, it, expect } from "vitest";
import {
  parseDefaultRunProperties,
  parseParagraphLevelProperties,
  parseListStyle,
} from "./text-style-parser.js";
import { ColorResolver } from "../color/color-resolver.js";

function createColorResolver() {
  return new ColorResolver(
    {
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
    },
    {
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
    },
  );
}

describe("parseDefaultRunProperties", () => {
  it("returns undefined for falsy input", () => {
    expect(parseDefaultRunProperties(undefined)).toBeUndefined();
    expect(parseDefaultRunProperties(null)).toBeUndefined();
  });

  it("parses fontSize from @_sz", () => {
    const result = parseDefaultRunProperties({ "@_sz": "1800" });
    expect(result?.fontSize).toBe(18);
  });

  it("parses fontFamily and fontFamilyEa", () => {
    const result = parseDefaultRunProperties({
      latin: { "@_typeface": "Arial" },
      ea: { "@_typeface": "MS Gothic" },
    });
    expect(result?.fontFamily).toBe("Arial");
    expect(result?.fontFamilyEa).toBe("MS Gothic");
  });

  it("parses bold, italic, underline, strikethrough", () => {
    const result = parseDefaultRunProperties({
      "@_b": "1",
      "@_i": "true",
      "@_u": "sng",
      "@_strike": "sngStrike",
    });
    expect(result?.bold).toBe(true);
    expect(result?.italic).toBe(true);
    expect(result?.underline).toBe(true);
    expect(result?.strikethrough).toBe(true);
  });

  it("parses false-like values", () => {
    const result = parseDefaultRunProperties({
      "@_b": "0",
      "@_u": "none",
      "@_strike": "noStrike",
    });
    expect(result?.bold).toBe(false);
    expect(result?.underline).toBe(false);
    expect(result?.strikethrough).toBe(false);
  });

  it("returns undefined for empty object", () => {
    expect(parseDefaultRunProperties({})).toBeUndefined();
  });

  it("parses solidFill with schemeClr when colorResolver is provided", () => {
    const result = parseDefaultRunProperties(
      {
        "@_sz": "1800",
        solidFill: { schemeClr: { "@_val": "lt1" } },
      },
      createColorResolver(),
    );
    expect(result?.fontSize).toBe(18);
    expect(result?.color).toEqual({ hex: "#FFFFFF", alpha: 1 });
  });

  it("parses solidFill with srgbClr when colorResolver is provided", () => {
    const result = parseDefaultRunProperties(
      {
        solidFill: { srgbClr: { "@_val": "FF0000" } },
      },
      createColorResolver(),
    );
    expect(result?.color).toEqual({ hex: "#FF0000", alpha: 1 });
  });

  it("does not parse color when colorResolver is not provided", () => {
    const result = parseDefaultRunProperties({
      "@_sz": "1800",
      solidFill: { schemeClr: { "@_val": "lt1" } },
    });
    expect(result?.color).toBeUndefined();
  });
});

describe("parseParagraphLevelProperties", () => {
  it("returns undefined for falsy input", () => {
    expect(parseParagraphLevelProperties(undefined)).toBeUndefined();
  });

  it("parses alignment, marginLeft, indent", () => {
    const result = parseParagraphLevelProperties({
      "@_algn": "ctr",
      "@_marL": "457200",
      "@_indent": "-228600",
    });
    expect(result?.alignment).toBe("ctr");
    expect(result?.marginLeft).toBe(457200);
    expect(result?.indent).toBe(-228600);
  });

  it("parses nested defRPr", () => {
    const result = parseParagraphLevelProperties({
      "@_algn": "l",
      defRPr: { "@_sz": "2400", "@_b": "1" },
    });
    expect(result?.alignment).toBe("l");
    expect(result?.defaultRunProperties?.fontSize).toBe(24);
    expect(result?.defaultRunProperties?.bold).toBe(true);
  });

  it("returns undefined for empty object", () => {
    expect(parseParagraphLevelProperties({})).toBeUndefined();
  });
});

describe("parseListStyle", () => {
  it("returns undefined for falsy input", () => {
    expect(parseListStyle(undefined)).toBeUndefined();
    expect(parseListStyle(null)).toBeUndefined();
  });

  it("parses defPPr and level properties", () => {
    const result = parseListStyle({
      defPPr: { "@_algn": "ctr" },
      lvl1pPr: { "@_marL": "0", defRPr: { "@_sz": "1800" } },
      lvl2pPr: { "@_marL": "457200", defRPr: { "@_sz": "1400" } },
    });

    expect(result).toBeDefined();
    expect(result!.defaultParagraph?.alignment).toBe("ctr");
    expect(result!.levels[0]?.marginLeft).toBe(0);
    expect(result!.levels[0]?.defaultRunProperties?.fontSize).toBe(18);
    expect(result!.levels[1]?.marginLeft).toBe(457200);
    expect(result!.levels[1]?.defaultRunProperties?.fontSize).toBe(14);
    for (let i = 2; i < 9; i++) {
      expect(result!.levels[i]).toBeUndefined();
    }
  });

  it("returns undefined when all levels are empty", () => {
    expect(parseListStyle({})).toBeUndefined();
  });

  it("returns style with only defPPr", () => {
    const result = parseListStyle({ defPPr: { "@_algn": "r" } });
    expect(result).toBeDefined();
    expect(result!.defaultParagraph?.alignment).toBe("r");
    expect(result!.levels.every((l) => l === undefined)).toBe(true);
  });

  it("returns style with only lvl1pPr", () => {
    const result = parseListStyle({
      lvl1pPr: { defRPr: { "@_sz": "3200" } },
    });
    expect(result).toBeDefined();
    expect(result!.defaultParagraph).toBeUndefined();
    expect(result!.levels[0]?.defaultRunProperties?.fontSize).toBe(32);
  });
});
