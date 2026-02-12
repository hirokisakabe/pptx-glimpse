import { describe, it, expect } from "vitest";
import { buildTransformAttr } from "./transform.js";
import type { Transform } from "../model/shape.js";

function makeTransform(overrides: Partial<Transform> = {}): Transform {
  return {
    offsetX: 0,
    offsetY: 0,
    extentWidth: 914400,
    extentHeight: 914400,
    rotation: 0,
    flipH: false,
    flipV: false,
    ...overrides,
  };
}

describe("buildTransformAttr", () => {
  it("generates translate only when no rotation or flip", () => {
    const t = makeTransform({ offsetX: 914400, offsetY: 914400 });
    expect(buildTransformAttr(t)).toBe("translate(96, 96)");
  });

  it("generates translate(0, 0) for zero offsets", () => {
    const t = makeTransform();
    expect(buildTransformAttr(t)).toBe("translate(0, 0)");
  });

  it("generates translate + rotate", () => {
    const t = makeTransform({
      offsetX: 914400,
      offsetY: 914400,
      extentWidth: 1828800, // 192px
      extentHeight: 914400, // 96px
      rotation: 45,
    });
    expect(buildTransformAttr(t)).toBe("translate(96, 96) rotate(45, 96, 48)");
  });

  it("calculates rotate center as width/2, height/2", () => {
    const t = makeTransform({
      extentWidth: 1828800, // 192px
      extentHeight: 1828800, // 192px
      rotation: 90,
    });
    expect(buildTransformAttr(t)).toBe("translate(0, 0) rotate(90, 96, 96)");
  });

  it("omits rotate when rotation is 0", () => {
    const t = makeTransform({ rotation: 0 });
    expect(buildTransformAttr(t)).not.toContain("rotate");
  });

  it("handles negative rotation", () => {
    const t = makeTransform({ rotation: -45 });
    expect(buildTransformAttr(t)).toContain("rotate(-45, 48, 48)");
  });

  it("handles 360 degree rotation", () => {
    const t = makeTransform({ rotation: 360 });
    expect(buildTransformAttr(t)).toContain("rotate(360, 48, 48)");
  });

  // --- flipH ---

  it("generates flipH transform", () => {
    const t = makeTransform({
      extentWidth: 1828800, // 192px
      extentHeight: 914400, // 96px
      flipH: true,
    });
    const result = buildTransformAttr(t);
    expect(result).toContain("translate(192, 0)");
    expect(result).toContain("scale(-1, 1)");
  });

  // --- flipV ---

  it("generates flipV transform", () => {
    const t = makeTransform({
      extentWidth: 1828800, // 192px
      extentHeight: 914400, // 96px
      flipV: true,
    });
    const result = buildTransformAttr(t);
    expect(result).toContain("translate(0, 96)");
    expect(result).toContain("scale(1, -1)");
  });

  // --- flipH + flipV ---

  it("generates both flips", () => {
    const t = makeTransform({
      extentWidth: 1828800, // 192px
      extentHeight: 914400, // 96px
      flipH: true,
      flipV: true,
    });
    const result = buildTransformAttr(t);
    expect(result).toContain("translate(192, 96)");
    expect(result).toContain("scale(-1, -1)");
  });

  // --- rotate + flip ---

  it("generates rotation and flipH", () => {
    const t = makeTransform({
      offsetX: 914400,
      offsetY: 914400,
      extentWidth: 1828800, // 192px
      extentHeight: 914400, // 96px
      rotation: 45,
      flipH: true,
    });
    const result = buildTransformAttr(t);
    expect(result).toBe("translate(96, 96) rotate(45, 96, 48) translate(192, 0) scale(-1, 1)");
  });

  it("generates rotation and flipV", () => {
    const t = makeTransform({
      extentWidth: 1828800,
      extentHeight: 914400,
      rotation: 90,
      flipV: true,
    });
    const result = buildTransformAttr(t);
    expect(result).toBe("translate(0, 0) rotate(90, 96, 48) translate(0, 96) scale(1, -1)");
  });

  it("generates rotation and both flips", () => {
    const t = makeTransform({
      extentWidth: 1828800,
      extentHeight: 914400,
      rotation: 180,
      flipH: true,
      flipV: true,
    });
    const result = buildTransformAttr(t);
    expect(result).toBe("translate(0, 0) rotate(180, 96, 48) translate(192, 96) scale(-1, -1)");
  });
});
