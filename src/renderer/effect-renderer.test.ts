import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderEffects } from "./effect-renderer.js";
import type { EffectList } from "../model/effect.js";

beforeEach(() => {
  let counter = 0;
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    return `test-uuid-${counter++}` as ReturnType<typeof crypto.randomUUID>;
  });
});

function makeEffects(overrides: Partial<EffectList> = {}): EffectList {
  return {
    outerShadow: null,
    innerShadow: null,
    glow: null,
    softEdge: null,
    ...overrides,
  };
}

describe("renderEffects", () => {
  it("returns empty result for null effects", () => {
    const result = renderEffects(null);
    expect(result.filterAttr).toBe("");
    expect(result.filterDefs).toBe("");
  });

  it("returns empty result when all effects are null", () => {
    const result = renderEffects(makeEffects());
    expect(result.filterAttr).toBe("");
    expect(result.filterDefs).toBe("");
  });

  // --- softEdge ---

  it("renders softEdge filter", () => {
    const result = renderEffects(makeEffects({ softEdge: { radius: 63500 } }));

    expect(result.filterAttr).toBe('filter="url(#effect-test-uuid-0)"');
    expect(result.filterDefs).toContain('<filter id="effect-test-uuid-0"');
    expect(result.filterDefs).toContain('<feGaussianBlur in="SourceAlpha"');
    expect(result.filterDefs).toContain('result="softEdgeMask"');
    expect(result.filterDefs).toContain(
      '<feComposite in="SourceGraphic" in2="softEdgeMask" operator="in"',
    );
  });

  // --- glow ---

  it("renders glow filter", () => {
    const result = renderEffects(
      makeEffects({
        glow: { radius: 127000, color: { hex: "#FF0000", alpha: 0.8 } },
      }),
    );

    expect(result.filterDefs).toContain('<feGaussianBlur in="SourceAlpha"');
    expect(result.filterDefs).toContain('result="glowBlur"');
    expect(result.filterDefs).toContain('flood-color="#FF0000"');
    expect(result.filterDefs).toContain('flood-opacity="0.8"');
    expect(result.filterDefs).toContain('result="glowFinal"');
    expect(result.filterDefs).toContain('<feMerge result="glowMerge">');
    expect(result.filterDefs).toContain('<feMergeNode in="glowFinal"/>');
    expect(result.filterDefs).toContain('<feMergeNode in="SourceGraphic"/>');
  });

  // --- outerShadow ---

  it("renders outerShadow filter", () => {
    const result = renderEffects(
      makeEffects({
        outerShadow: {
          blurRadius: 50800,
          distance: 38100,
          direction: 45,
          color: { hex: "#000000", alpha: 0.5 },
          alignment: "br",
          rotateWithShape: true,
        },
      }),
    );

    expect(result.filterDefs).toContain('<feGaussianBlur in="SourceAlpha"');
    expect(result.filterDefs).toContain('result="shadowBlur"');
    expect(result.filterDefs).toContain("<feOffset");
    expect(result.filterDefs).toContain('result="shadowOffset"');
    expect(result.filterDefs).toContain('flood-color="#000000"');
    expect(result.filterDefs).toContain('flood-opacity="0.5"');
    expect(result.filterDefs).toContain('<feMerge result="outerShadowMerge">');
    expect(result.filterDefs).toContain('<feMergeNode in="shadowFinal"/>');
    expect(result.filterDefs).toContain('<feMergeNode in="SourceGraphic"/>');
  });

  it("calculates outerShadow dx/dy for 0 degrees (right)", () => {
    const result = renderEffects(
      makeEffects({
        outerShadow: {
          blurRadius: 0,
          distance: 914400, // 96px
          direction: 0,
          color: { hex: "#000000", alpha: 1 },
          alignment: "b",
          rotateWithShape: false,
        },
      }),
    );
    // 0 deg → cos(0)=1, sin(0)=0 → dx=96, dy=0
    expect(result.filterDefs).toContain('dx="96"');
    expect(result.filterDefs).toContain('dy="0"');
  });

  it("calculates outerShadow dx/dy for 90 degrees (down)", () => {
    const result = renderEffects(
      makeEffects({
        outerShadow: {
          blurRadius: 0,
          distance: 914400, // 96px
          direction: 90,
          color: { hex: "#000000", alpha: 1 },
          alignment: "b",
          rotateWithShape: false,
        },
      }),
    );
    // 90 deg → cos(90)≈0, sin(90)=1 → dx≈0, dy=96
    expect(result.filterDefs).toContain('dx="0"');
    expect(result.filterDefs).toContain('dy="96"');
  });

  it("calculates outerShadow dx/dy for 180 degrees (left)", () => {
    const result = renderEffects(
      makeEffects({
        outerShadow: {
          blurRadius: 0,
          distance: 914400,
          direction: 180,
          color: { hex: "#000000", alpha: 1 },
          alignment: "b",
          rotateWithShape: false,
        },
      }),
    );
    // 180 deg → cos(180)=-1, sin(180)≈0 → dx=-96, dy≈0
    expect(result.filterDefs).toContain('dx="-96"');
    expect(result.filterDefs).toContain('dy="0"');
  });

  // --- innerShadow ---

  it("renders innerShadow filter", () => {
    const result = renderEffects(
      makeEffects({
        innerShadow: {
          blurRadius: 63500,
          distance: 38100,
          direction: 135,
          color: { hex: "#FF0000", alpha: 0.75 },
        },
      }),
    );

    expect(result.filterDefs).toContain('<feComponentTransfer in="SourceAlpha"');
    expect(result.filterDefs).toContain('<feFuncA type="table" tableValues="1 0"/>');
    expect(result.filterDefs).toContain("<feGaussianBlur");
    expect(result.filterDefs).toContain("<feOffset");
    expect(result.filterDefs).toContain('flood-color="#FF0000"');
    expect(result.filterDefs).toContain('flood-opacity="0.75"');
    expect(result.filterDefs).toContain('operator="in"');
    expect(result.filterDefs).toContain('operator="over"');
  });

  // --- filter 構造 ---

  it("generates filter with correct structure", () => {
    const result = renderEffects(makeEffects({ softEdge: { radius: 63500 } }));

    expect(result.filterDefs).toContain('x="-50%"');
    expect(result.filterDefs).toContain('y="-50%"');
    expect(result.filterDefs).toContain('width="200%"');
    expect(result.filterDefs).toContain('height="200%"');
    expect(result.filterDefs).toContain('color-interpolation-filters="sRGB"');
    expect(result.filterDefs).toContain("</filter>");
  });

  // --- 複合エフェクト ---

  it("chains softEdge into glow", () => {
    const result = renderEffects(
      makeEffects({
        glow: { radius: 127000, color: { hex: "#00FF00", alpha: 1 } },
        softEdge: { radius: 63500 },
      }),
    );

    // softEdge が先に適用される
    expect(result.filterDefs).toContain('result="softEdgeResult"');
    // glow は softEdgeResult を入力として使用
    expect(result.filterDefs).toContain('in="softEdgeResult"');
    expect(result.filterDefs).toContain('<feMergeNode in="softEdgeResult"/>');
  });

  it("renders all four effects together", () => {
    const result = renderEffects({
      outerShadow: {
        blurRadius: 50800,
        distance: 38100,
        direction: 45,
        color: { hex: "#000000", alpha: 0.5 },
        alignment: "b",
        rotateWithShape: false,
      },
      innerShadow: {
        blurRadius: 63500,
        distance: 38100,
        direction: 135,
        color: { hex: "#FF0000", alpha: 0.75 },
      },
      glow: { radius: 127000, color: { hex: "#00FF00", alpha: 1 } },
      softEdge: { radius: 63500 },
    });

    expect(result.filterDefs).toContain("softEdgeMask");
    expect(result.filterDefs).toContain("glowMerge");
    expect(result.filterDefs).toContain("outerShadowMerge");
    expect(result.filterDefs).toContain("innerShdw");
    expect(result.filterAttr).toContain("filter=");
  });
});
