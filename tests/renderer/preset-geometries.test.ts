import { describe, it, expect } from "vitest";
import { getPresetGeometrySvg } from "../../src/renderer/geometry/preset-geometries.js";

describe("getPresetGeometrySvg", () => {
  // --- Basic shapes ---

  it("generates rect", () => {
    const svg = getPresetGeometrySvg("rect", 100, 50, {});
    expect(svg).toBe('<rect width="100" height="50"/>');
  });

  it("generates ellipse", () => {
    const svg = getPresetGeometrySvg("ellipse", 200, 100, {});
    expect(svg).toContain("<ellipse");
    expect(svg).toContain('cx="100"');
    expect(svg).toContain('rx="100"');
    expect(svg).toContain('ry="50"');
  });

  it("generates roundRect with default radius", () => {
    const svg = getPresetGeometrySvg("roundRect", 100, 100, {});
    expect(svg).toContain("<rect");
    expect(svg).toContain("rx=");
    expect(svg).toContain("ry=");
  });

  it("generates roundRect with custom adjust value", () => {
    const svg = getPresetGeometrySvg("roundRect", 100, 100, { adj: 50000 });
    expect(svg).toContain("rx=");
  });

  it("generates triangle", () => {
    const svg = getPresetGeometrySvg("triangle", 100, 80, {});
    expect(svg).toContain("<polygon");
    expect(svg).toContain("points=");
  });

  it("generates diamond", () => {
    const svg = getPresetGeometrySvg("diamond", 100, 100, {});
    expect(svg).toContain("<polygon");
  });

  it("generates rightArrow", () => {
    const svg = getPresetGeometrySvg("rightArrow", 200, 100, {});
    expect(svg).toContain("<polygon");
  });

  it("falls back to rect for unknown preset", () => {
    const svg = getPresetGeometrySvg("unknownShape", 100, 50, {});
    expect(svg).toBe('<rect width="100" height="50"/>');
  });

  // --- Additional polygons ---

  it.each(["heptagon", "octagon", "decagon", "dodecagon"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 100, 100, {});
    expect(svg).toContain("<polygon");
    expect(svg).toContain("points=");
  });

  // --- Stars ---

  it.each(["star6", "star8", "star10", "star12", "star16", "star24", "star32"])(
    "generates %s",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 100, 100, {});
      expect(svg).toContain("<polygon");
      expect(svg).toContain("points=");
    },
  );

  it.each(["irregularSeal1", "irregularSeal2"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 100, 100, {});
    expect(svg).toContain("<polygon");
  });

  // --- Additional arrows ---

  it.each(["leftRightArrow", "upDownArrow"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 200, 100, {});
    expect(svg).toContain("<polygon");
  });

  it.each(["notchedRightArrow"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 200, 100, {});
    expect(svg).toContain("<polygon");
  });

  it.each(["stripedRightArrow", "quadArrow", "leftUpArrow", "leftRightUpArrow"])(
    "generates %s as path",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 200, 100, {});
      expect(svg).toContain("<path");
    },
  );

  it.each(["chevron", "homePlate"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 200, 100, {});
    expect(svg).toContain("<polygon");
  });

  it.each(["bentArrow", "bendUpArrow"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 200, 100, {});
    expect(svg).toContain("<polygon");
  });

  it("generates uturnArrow", () => {
    const svg = getPresetGeometrySvg("uturnArrow", 200, 200, {});
    expect(svg).toContain("<path");
  });

  // --- Flowchart shapes ---

  it("generates flowChartProcess as rect", () => {
    const svg = getPresetGeometrySvg("flowChartProcess", 100, 50, {});
    expect(svg).toBe('<rect width="100" height="50"/>');
  });

  it("generates flowChartAlternateProcess with rounded corners", () => {
    const svg = getPresetGeometrySvg("flowChartAlternateProcess", 120, 60, {});
    expect(svg).toContain("<rect");
    expect(svg).toContain("rx=");
  });

  it("generates flowChartDecision as diamond", () => {
    const svg = getPresetGeometrySvg("flowChartDecision", 100, 100, {});
    expect(svg).toContain("<polygon");
    expect(svg).toContain("50,0");
  });

  it("generates flowChartInputOutput as parallelogram", () => {
    const svg = getPresetGeometrySvg("flowChartInputOutput", 100, 50, {});
    expect(svg).toContain("<polygon");
  });

  it.each([
    "flowChartPredefinedProcess",
    "flowChartInternalStorage",
    "flowChartDocument",
    "flowChartMultidocument",
    "flowChartPunchedTape",
    "flowChartOnlineStorage",
    "flowChartDelay",
    "flowChartDisplay",
    "flowChartMagneticTape",
    "flowChartMagneticDisk",
    "flowChartMagneticDrum",
    "flowChartSummingJunction",
    "flowChartOr",
    "flowChartSort",
  ])("generates %s as path", (shape) => {
    const svg = getPresetGeometrySvg(shape, 100, 100, {});
    expect(svg).toContain("<path");
  });

  it("generates flowChartTerminator with full rounded ends", () => {
    const svg = getPresetGeometrySvg("flowChartTerminator", 200, 60, {});
    expect(svg).toContain("<rect");
    expect(svg).toContain('rx="30"');
  });

  it.each([
    "flowChartPreparation",
    "flowChartManualInput",
    "flowChartManualOperation",
    "flowChartOffpageConnector",
    "flowChartPunchedCard",
    "flowChartCollate",
    "flowChartExtract",
    "flowChartMerge",
  ])("generates %s as polygon", (shape) => {
    const svg = getPresetGeometrySvg(shape, 100, 100, {});
    expect(svg).toContain("<polygon");
  });

  it("generates flowChartConnector as ellipse", () => {
    const svg = getPresetGeometrySvg("flowChartConnector", 100, 100, {});
    expect(svg).toContain("<ellipse");
  });

  // --- Callout shapes ---

  it.each(["wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout"])(
    "generates %s with default callout tip",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 200, 100, {});
      expect(svg).toContain("<path");
    },
  );

  it("generates wedgeRectCallout with custom adj values", () => {
    const svg = getPresetGeometrySvg("wedgeRectCallout", 200, 100, {
      adj1: 0,
      adj2: 100000,
    });
    expect(svg).toContain("<path");
  });

  it.each(["borderCallout1", "borderCallout2", "borderCallout3"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 200, 100, {});
    expect(svg).toContain("<path");
  });

  // --- Arc shapes ---

  it.each(["arc", "chord", "pie", "blockArc"])("generates %s", (shape) => {
    const svg = getPresetGeometrySvg(shape, 100, 100, {});
    expect(svg).toContain("<path");
  });

  it("generates pie with custom angles", () => {
    const svg = getPresetGeometrySvg("pie", 100, 100, {
      adj1: 5400000,
      adj2: 16200000,
    });
    expect(svg).toContain("<path");
    expect(svg).toContain("A 50 50");
  });

  // --- Math shapes ---

  it("generates mathPlus as polygon", () => {
    const svg = getPresetGeometrySvg("mathPlus", 100, 100, {});
    expect(svg).toContain("<polygon");
  });

  it("generates mathMinus as rect", () => {
    const svg = getPresetGeometrySvg("mathMinus", 100, 100, {});
    expect(svg).toContain("<rect");
  });

  it.each(["mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"])(
    "generates %s as path",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 100, 100, {});
      expect(svg).toContain("<path");
    },
  );

  // --- Other shapes ---

  it("generates plus as polygon", () => {
    const svg = getPresetGeometrySvg("plus", 100, 100, {});
    expect(svg).toContain("<polygon");
  });

  it.each(["corner", "halfFrame", "snip1Rect", "snip2SameRect", "snip2DiagRect"])(
    "generates %s as polygon",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 100, 100, {});
      expect(svg).toMatch(/<polygon|<path/);
    },
  );

  it.each([
    "foldedCorner",
    "plaque",
    "can",
    "cube",
    "snipRoundRect",
    "round1Rect",
    "round2SameRect",
    "round2DiagRect",
  ])("generates %s as path", (shape) => {
    const svg = getPresetGeometrySvg(shape, 100, 100, {});
    expect(svg).toContain("<path");
  });

  it("generates donut with fill-rule evenodd", () => {
    const svg = getPresetGeometrySvg("donut", 100, 100, {});
    expect(svg).toContain('fill-rule="evenodd"');
    expect(svg).toContain("<path");
  });

  it("generates frame with fill-rule evenodd", () => {
    const svg = getPresetGeometrySvg("frame", 100, 100, {});
    expect(svg).toContain('fill-rule="evenodd"');
  });

  it("generates noSmoking with fill-rule evenodd", () => {
    const svg = getPresetGeometrySvg("noSmoking", 100, 100, {});
    expect(svg).toContain('fill-rule="evenodd"');
  });

  it.each(["smileyFace", "bevel", "lightningBolt", "moon", "teardrop", "sun"])(
    "generates %s",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 100, 100, {});
      expect(svg).toMatch(/<path|<polygon/);
    },
  );

  // --- Brackets and braces ---

  it.each(["leftBracket", "rightBracket", "leftBrace", "rightBrace", "bracketPair", "bracePair"])(
    "generates %s as path",
    (shape) => {
      const svg = getPresetGeometrySvg(shape, 20, 100, {});
      expect(svg).toContain("<path");
    },
  );

  // --- Banners ---

  it.each(["wave", "doubleWave", "ribbon", "ribbon2"])("generates %s as path", (shape) => {
    const svg = getPresetGeometrySvg(shape, 200, 100, {});
    expect(svg).toContain("<path");
  });

  it("generates diagStripe as polygon", () => {
    const svg = getPresetGeometrySvg("diagStripe", 100, 100, {});
    expect(svg).toContain("<polygon");
  });
});
