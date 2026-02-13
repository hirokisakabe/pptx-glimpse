import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderBlipEffects } from "./blip-effect-renderer.js";
import type { BlipEffects } from "../model/effect.js";

beforeEach(() => {
  let counter = 0;
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    return `test-uuid-${counter++}` as ReturnType<typeof crypto.randomUUID>;
  });
});

function makeBlipEffects(overrides: Partial<BlipEffects> = {}): BlipEffects {
  return {
    grayscale: false,
    biLevel: null,
    blur: null,
    lum: null,
    duotone: null,
    ...overrides,
  };
}

describe("renderBlipEffects", () => {
  it("returns empty for null input", () => {
    const result = renderBlipEffects(null);
    expect(result.filterAttr).toBe("");
    expect(result.filterDefs).toBe("");
  });

  it("returns empty when no effects are active", () => {
    const result = renderBlipEffects(makeBlipEffects());
    expect(result.filterAttr).toBe("");
    expect(result.filterDefs).toBe("");
  });

  it("renders grayscale filter", () => {
    const result = renderBlipEffects(makeBlipEffects({ grayscale: true }));
    expect(result.filterDefs).toContain('<filter id="blip-effect-test-uuid-0"');
    expect(result.filterDefs).toContain('type="saturate" values="0"');
    expect(result.filterAttr).toBe('filter="url(#blip-effect-test-uuid-0)"');
  });

  it("renders biLevel filter with threshold", () => {
    const result = renderBlipEffects(makeBlipEffects({ biLevel: { threshold: 0.5 } }));
    expect(result.filterDefs).toContain('type="saturate" values="0"');
    expect(result.filterDefs).toContain("<feComponentTransfer");
    expect(result.filterDefs).toContain('type="discrete"');
  });

  it("renders blur filter", () => {
    const result = renderBlipEffects(makeBlipEffects({ blur: { radius: 50800, grow: false } }));
    expect(result.filterDefs).toContain("<feGaussianBlur");
    expect(result.filterDefs).toContain("stdDeviation=");
  });

  it("renders lum filter (brightness/contrast)", () => {
    const result = renderBlipEffects(makeBlipEffects({ lum: { brightness: 0.7, contrast: -0.3 } }));
    expect(result.filterDefs).toContain("<feComponentTransfer");
    expect(result.filterDefs).toContain('type="linear"');
    expect(result.filterDefs).toContain("slope=");
    expect(result.filterDefs).toContain("intercept=");
  });

  it("renders duotone filter", () => {
    const result = renderBlipEffects(
      makeBlipEffects({
        duotone: {
          color1: { hex: "#000000", alpha: 1 },
          color2: { hex: "#D9C3A5", alpha: 1 },
        },
      }),
    );
    expect(result.filterDefs).toContain('type="saturate" values="0"');
    expect(result.filterDefs).toContain("<feComponentTransfer");
    expect(result.filterDefs).toContain('type="table"');
    expect(result.filterDefs).toContain("tableValues=");
  });

  it("does not duplicate grayscale when both grayscale and duotone are set", () => {
    const result = renderBlipEffects(
      makeBlipEffects({
        grayscale: true,
        duotone: {
          color1: { hex: "#000000", alpha: 1 },
          color2: { hex: "#FFFFFF", alpha: 1 },
        },
      }),
    );
    const matches = result.filterDefs.match(/type="saturate" values="0"/g);
    expect(matches).toHaveLength(1);
  });

  it("chains multiple effects", () => {
    const result = renderBlipEffects(
      makeBlipEffects({
        grayscale: true,
        blur: { radius: 25400, grow: false },
        lum: { brightness: 0.5, contrast: 0 },
      }),
    );
    expect(result.filterDefs).toContain("feColorMatrix");
    expect(result.filterDefs).toContain("feGaussianBlur");
    expect(result.filterDefs).toContain("feComponentTransfer");
  });
});
