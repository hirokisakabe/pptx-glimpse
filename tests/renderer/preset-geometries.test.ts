import { describe, it, expect } from "vitest";
import { getPresetGeometrySvg } from "../../src/renderer/geometry/preset-geometries.js";

describe("getPresetGeometrySvg", () => {
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
});
