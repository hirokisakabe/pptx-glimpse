import { describe, it, expect } from "vitest";
import { applyColorTransforms } from "./color-transforms.js";

describe("applyColorTransforms", () => {
  it("returns unchanged color when no transforms are applied", () => {
    const result = applyColorTransforms({ hex: "#FF0000", alpha: 1 }, {});
    expect(result).toEqual({ hex: "#FF0000", alpha: 1 });
  });

  // --- lumMod / lumOff ---

  it("applies lumMod to darken color (50%)", () => {
    // #808080 → HSL(0, 0, 0.502) → lumMod 50% → l=0.251 → ~#404040
    const result = applyColorTransforms(
      { hex: "#808080", alpha: 1 },
      { lumMod: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#404040");
    expect(result.alpha).toBe(1);
  });

  it("applies lumMod 100% without change", () => {
    const result = applyColorTransforms(
      { hex: "#808080", alpha: 1 },
      { lumMod: { "@_val": "100000" } },
    );
    expect(result.hex).toBe("#808080");
  });

  it("applies lumOff to brighten black", () => {
    // #000000 → HSL(0, 0, 0) → lumOff +50% → l=0.5 → #808080
    const result = applyColorTransforms(
      { hex: "#000000", alpha: 1 },
      { lumOff: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#808080");
  });

  it("applies lumMod + lumOff together", () => {
    // #808080 → l=0.502 → lumMod 75% + lumOff 25% → l = 0.502*0.75 + 0.25 = 0.6265
    const result = applyColorTransforms(
      { hex: "#808080", alpha: 1 },
      { lumMod: { "@_val": "75000" }, lumOff: { "@_val": "25000" } },
    );
    // 結果が元より明るくなっている
    expect(result.hex).not.toBe("#808080");
    expect(result.alpha).toBe(1);
  });

  it("clamps luminance to 1 when lumMod + lumOff exceeds 100%", () => {
    const result = applyColorTransforms(
      { hex: "#FFFFFF", alpha: 1 },
      { lumMod: { "@_val": "200000" }, lumOff: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#ffffff");
  });

  it("clamps luminance to 0 when lumMod is 0", () => {
    const result = applyColorTransforms({ hex: "#808080", alpha: 1 }, { lumMod: { "@_val": "0" } });
    expect(result.hex).toBe("#000000");
  });

  // --- tint (白方向ブレンド) ---

  it("applies tint 50% to red", () => {
    // #FF0000 → tint 50% → r=255+(255-255)*0.5=255, g=0+(255-0)*0.5=128, b=0+128=128
    const result = applyColorTransforms(
      { hex: "#FF0000", alpha: 1 },
      { tint: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#ff8080");
  });

  it("applies tint 100% to produce white", () => {
    const result = applyColorTransforms(
      { hex: "#000000", alpha: 1 },
      { tint: { "@_val": "100000" } },
    );
    expect(result.hex).toBe("#ffffff");
  });

  it("applies tint 0% without change", () => {
    const result = applyColorTransforms({ hex: "#FF0000", alpha: 1 }, { tint: { "@_val": "0" } });
    expect(result.hex).toBe("#ff0000");
  });

  it("applies tint to white (no change)", () => {
    const result = applyColorTransforms(
      { hex: "#FFFFFF", alpha: 1 },
      { tint: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#ffffff");
  });

  // --- shade (黒方向ブレンド) ---

  it("applies shade 50% to red", () => {
    // #FF0000 → shade 50% → r=255*0.5=128, g=0, b=0
    const result = applyColorTransforms(
      { hex: "#FF0000", alpha: 1 },
      { shade: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#800000");
  });

  it("applies shade 0% to produce black", () => {
    const result = applyColorTransforms({ hex: "#FFFFFF", alpha: 1 }, { shade: { "@_val": "0" } });
    expect(result.hex).toBe("#000000");
  });

  it("applies shade 100% without change", () => {
    const result = applyColorTransforms(
      { hex: "#FF0000", alpha: 1 },
      { shade: { "@_val": "100000" } },
    );
    expect(result.hex).toBe("#ff0000");
  });

  it("applies shade to black (no change)", () => {
    const result = applyColorTransforms(
      { hex: "#000000", alpha: 1 },
      { shade: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#000000");
  });

  // --- alpha ---

  it("applies alpha 50%", () => {
    const result = applyColorTransforms(
      { hex: "#FF0000", alpha: 1 },
      { alpha: { "@_val": "50000" } },
    );
    expect(result.hex).toBe("#FF0000");
    expect(result.alpha).toBe(0.5);
  });

  it("applies alpha 0% (fully transparent)", () => {
    const result = applyColorTransforms({ hex: "#FF0000", alpha: 1 }, { alpha: { "@_val": "0" } });
    expect(result.alpha).toBe(0);
  });

  it("applies alpha 100%", () => {
    const result = applyColorTransforms(
      { hex: "#FF0000", alpha: 1 },
      { alpha: { "@_val": "100000" } },
    );
    expect(result.alpha).toBe(1);
  });

  // --- 複合変換 ---

  it("applies multiple transforms in sequence", () => {
    const result = applyColorTransforms(
      { hex: "#4472C4", alpha: 1 },
      {
        lumMod: { "@_val": "75000" },
        tint: { "@_val": "20000" },
        alpha: { "@_val": "80000" },
      },
    );
    expect(result.alpha).toBe(0.8);
    expect(result.hex).toBeTruthy();
  });

  // --- HSL往復変換の精度 ---

  it("preserves red through lumMod 100%", () => {
    const result = applyColorTransforms(
      { hex: "#FF0000", alpha: 1 },
      { lumMod: { "@_val": "100000" } },
    );
    expect(result.hex.toLowerCase()).toBe("#ff0000");
  });

  it("preserves green through lumMod 100%", () => {
    const result = applyColorTransforms(
      { hex: "#00FF00", alpha: 1 },
      { lumMod: { "@_val": "100000" } },
    );
    expect(result.hex.toLowerCase()).toBe("#00ff00");
  });

  it("preserves blue through lumMod 100%", () => {
    const result = applyColorTransforms(
      { hex: "#0000FF", alpha: 1 },
      { lumMod: { "@_val": "100000" } },
    );
    expect(result.hex.toLowerCase()).toBe("#0000ff");
  });

  it("preserves white through lumMod 100%", () => {
    const result = applyColorTransforms(
      { hex: "#FFFFFF", alpha: 1 },
      { lumMod: { "@_val": "100000" } },
    );
    expect(result.hex.toLowerCase()).toBe("#ffffff");
  });

  it("preserves black through lumMod 100%", () => {
    const result = applyColorTransforms(
      { hex: "#000000", alpha: 1 },
      { lumMod: { "@_val": "100000" } },
    );
    expect(result.hex.toLowerCase()).toBe("#000000");
  });

  it("handles lowercase hex input", () => {
    const result = applyColorTransforms({ hex: "#ff0000", alpha: 1 }, {});
    expect(result.hex).toBe("#ff0000");
  });
});
