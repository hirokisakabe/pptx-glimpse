import { describe, expect, it } from "vitest";

import {
  buildSourceGroupMappingMatrix,
  buildSourceTransformMatrix,
  composeSourceAncestorMatrices,
  decomposeSourceTransformMatrix,
  invertSourceAffineMatrix,
  multiplySourceAffineMatrices,
  type SourceAffineMatrix,
  type SourceTransformMatrixResult,
} from "./group-transform-matrix.js";
import type { SourceTransform } from "./shapes.js";
import { asEmu, asOoxmlAngle } from "./units.js";

function transform(
  x: number,
  y: number,
  width: number,
  height: number,
  options: Partial<Pick<SourceTransform, "rotation" | "flipHorizontal" | "flipVertical">> = {},
): SourceTransform {
  return {
    offsetX: asEmu(x),
    offsetY: asEmu(y),
    width: asEmu(width),
    height: asEmu(height),
    ...options,
  };
}

function valueOf<T>(result: SourceTransformMatrixResult<T>): T {
  if (!result.ok) throw new Error(`unexpected matrix rejection: ${result.reason}`);
  return result.value;
}

function expectMatrixClose(actual: SourceAffineMatrix, expected: SourceAffineMatrix): void {
  for (const key of ["a", "b", "c", "d", "e", "f"] as const) {
    expect(actual[key]).toBeCloseTo(expected[key], 8);
  }
}

describe("buildSourceGroupMappingMatrix", () => {
  it("builds identity and scaled/translated parent-local mappings from group and child transforms", () => {
    const identityTransform = transform(10, 20, 100, 40);
    expect(valueOf(buildSourceGroupMappingMatrix(identityTransform, identityTransform))).toEqual({
      a: 1,
      b: 0,
      c: 0,
      d: 1,
      e: 0,
      f: 0,
    });

    expect(
      valueOf(buildSourceGroupMappingMatrix(transform(10, 20, 200, 120), transform(5, 7, 100, 40))),
    ).toEqual({ a: 2, b: 0, c: 0, d: 3, e: 0, f: -1 });
  });

  it("applies rotation and horizontal/vertical flips around the group extent center", () => {
    const matrix = valueOf(
      buildSourceGroupMappingMatrix(
        transform(10, 20, 200, 100, {
          rotation: asOoxmlAngle(90 * 60000),
          flipHorizontal: true,
          flipVertical: true,
        }),
        transform(0, 0, 100, 50),
      ),
    );

    expectMatrixClose(matrix, { a: 0, b: -2, c: 2, d: 0, e: 60, f: 170 });
  });

  it.each([
    [undefined, transform(0, 0, 1, 1), "missing-transform"],
    [transform(0, 0, 1, 1), undefined, "missing-child-transform"],
    [transform(0, 0, 1, 1), transform(0, 0, 0, 1), "zero-child-extent"],
    [transform(0, 0, 1, 1), transform(0, 0, 1, 0), "zero-child-extent"],
    [transform(0, 0, -1, 1), transform(0, 0, 1, 1), "invalid-extent"],
    [transform(0, 0, 1, 1), transform(0, 0, -1, 1), "invalid-extent"],
    [transform(Number.NaN, 0, 1, 1), transform(0, 0, 1, 1), "non-finite-transform"],
  ] as const)("rejects an invalid group coordinate space (%s)", (group, child, reason) => {
    expect(buildSourceGroupMappingMatrix(group, child)).toEqual({ ok: false, reason });
  });
});

describe("source affine matrix composition", () => {
  it("composes nested ancestors outer-to-inner", () => {
    const outer = valueOf(
      buildSourceGroupMappingMatrix(transform(10, 20, 200, 100), transform(0, 0, 100, 50)),
    );
    const inner = valueOf(
      buildSourceGroupMappingMatrix(transform(5, 7, 30, 40), transform(1, 2, 10, 10)),
    );

    const composed = valueOf(composeSourceAncestorMatrices([outer, inner]));
    expectMatrixClose(composed, multiplySourceAffineMatrices(outer, inner));
    expectMatrixClose(composed, { a: 6, b: 0, c: 0, d: 8, e: 14, f: 18 });
  });

  it("inverts a reversible composed matrix", () => {
    const matrix = valueOf(
      buildSourceGroupMappingMatrix(
        transform(10, 20, 200, 100, {
          rotation: asOoxmlAngle(30 * 60000),
          flipVertical: true,
        }),
        transform(5, 7, 100, 25),
      ),
    );
    const inverse = valueOf(invertSourceAffineMatrix(matrix));
    expectMatrixClose(multiplySourceAffineMatrices(matrix, inverse), {
      a: 1,
      b: 0,
      c: 0,
      d: 1,
      e: 0,
      f: 0,
    });
  });

  it("rejects singular and non-finite matrices", () => {
    expect(invertSourceAffineMatrix({ a: 1, b: 2, c: 2, d: 4, e: 0, f: 0 })).toEqual({
      ok: false,
      reason: "singular-matrix",
    });
    expect(composeSourceAncestorMatrices([{ a: 1, b: 0, c: 0, d: 1, e: Infinity, f: 0 }])).toEqual({
      ok: false,
      reason: "non-finite-matrix",
    });
    expect(invertSourceAffineMatrix({ a: 1, b: 0, c: 0, d: 1, e: Number.NaN, f: 0 })).toEqual({
      ok: false,
      reason: "non-finite-matrix",
    });
    expect(
      composeSourceAncestorMatrices([
        { a: 1e308, b: 0, c: 0, d: 1e308, e: 0, f: 0 },
        { a: 1e308, b: 0, c: 0, d: 1e308, e: 0, f: 0 },
      ]),
    ).toEqual({ ok: false, reason: "non-finite-matrix" });
  });

  it("inverts small but well-conditioned scales without treating magnitude as singularity", () => {
    const tinyScale = { a: 1e-8, b: 0, c: 0, d: 2e-8, e: 3, f: -4 };
    const inverse = valueOf(invertSourceAffineMatrix(tinyScale));

    expectMatrixClose(multiplySourceAffineMatrices(tinyScale, inverse), {
      a: 1,
      b: 0,
      c: 0,
      d: 1,
      e: 0,
      f: 0,
    });
  });

  it("avoids intermediate overflow while inverting large finite coefficients", () => {
    const largeRotationScale = {
      a: 1e308,
      b: -1e308,
      c: 1e308,
      d: 1e308,
      e: 0,
      f: 0,
    };
    const inverse = valueOf(invertSourceAffineMatrix(largeRotationScale));

    expectMatrixClose(multiplySourceAffineMatrices(largeRotationScale, inverse), {
      a: 1,
      b: 0,
      c: 0,
      d: 1,
      e: 0,
      f: 0,
    });
  });
});

describe("source transform decomposition and reconstruction", () => {
  it.each([
    ["translation", transform(12, -8, 100, 100)],
    ["uniform scale", transform(12, -8, 240, 240)],
    ["non-uniform scale", transform(12, -8, 240, 75)],
    ["rotation", transform(12, -8, 240, 75, { rotation: asOoxmlAngle(33 * 60000) })],
    ["horizontal flip", transform(12, -8, 240, 75, { flipHorizontal: true })],
    ["vertical flip", transform(12, -8, 240, 75, { flipVertical: true })],
    [
      "rotation, flip, and non-uniform scale",
      transform(12, -8, 240, 75, {
        rotation: asOoxmlAngle(-47 * 60000),
        flipVertical: true,
      }),
    ],
  ] as const)("round-trips a representable %s transform through OOXML values", (_, source) => {
    const original = valueOf(buildSourceTransformMatrix(source));
    const decomposed = valueOf(decomposeSourceTransformMatrix(original));
    const reconstructed = valueOf(buildSourceTransformMatrix(decomposed));

    expectMatrixClose(reconstructed, original);
    expect(Number.isInteger(Number(decomposed.offsetX))).toBe(true);
    expect(Number.isInteger(Number(decomposed.offsetY))).toBe(true);
    expect(Number.isInteger(Number(decomposed.width))).toBe(true);
    expect(Number.isInteger(Number(decomposed.height))).toBe(true);
    expect(Number.isInteger(Number(decomposed.rotation))).toBe(true);
  });

  it("rejects shear, singular/non-finite matrices, and quantization mismatch", () => {
    expect(decomposeSourceTransformMatrix({ a: 1, b: 0, c: 0.25, d: 1, e: 0, f: 0 })).toEqual({
      ok: false,
      reason: "shear",
    });
    expect(decomposeSourceTransformMatrix({ a: 0, b: 0, c: 0, d: 1, e: 0, f: 0 })).toEqual({
      ok: false,
      reason: "singular-matrix",
    });
    expect(decomposeSourceTransformMatrix({ a: 1, b: 0, c: 0, d: 1, e: Infinity, f: 0 })).toEqual({
      ok: false,
      reason: "non-finite-matrix",
    });
    expect(decomposeSourceTransformMatrix({ a: 100, b: 0, c: 0, d: 100, e: 0.25, f: 0 })).toEqual({
      ok: false,
      reason: "quantization-mismatch",
    });
    expect(
      decomposeSourceTransformMatrix({
        a: 1_000_000_000_000.25,
        b: 0,
        c: 0,
        d: 100,
        e: 0,
        f: 0,
      }),
    ).toEqual({ ok: false, reason: "quantization-mismatch" });
    expect(
      decomposeSourceTransformMatrix({
        a: 100,
        b: 0,
        c: 0,
        d: 100,
        e: 1_000_000_000_000.25,
        f: 0,
      }),
    ).toEqual({ ok: false, reason: "quantization-mismatch" });
  });

  it("rejects zero, negative, and non-finite OOXML extents", () => {
    expect(buildSourceTransformMatrix(transform(0, 0, 0, 1))).toEqual({
      ok: false,
      reason: "invalid-extent",
    });
    expect(buildSourceTransformMatrix(transform(0, 0, 1, -1))).toEqual({
      ok: false,
      reason: "invalid-extent",
    });
    expect(buildSourceTransformMatrix(transform(0, 0, Infinity, 1))).toEqual({
      ok: false,
      reason: "non-finite-transform",
    });
  });

  it("allows bounded floating-point error after ancestor composition and inversion", () => {
    const parent = valueOf(
      buildSourceGroupMappingMatrix(
        transform(1_000_000, 2_000_000, 9_144_000, 5_143_500, {
          rotation: asOoxmlAngle(37 * 60000),
          flipHorizontal: true,
        }),
        transform(123_456, 789_012, 3_000_000, 7_000_000),
      ),
    );
    const local = valueOf(
      buildSourceTransformMatrix(
        transform(2_345_678, 1_234_567, 4_567_890, 2_345_678, {
          rotation: asOoxmlAngle(-23 * 60000),
          flipVertical: true,
        }),
      ),
    );
    const inverseParent = valueOf(invertSourceAffineMatrix(parent));
    const roundTripped = multiplySourceAffineMatrices(
      inverseParent,
      multiplySourceAffineMatrices(parent, local),
    );

    expect(decomposeSourceTransformMatrix(roundTripped).ok).toBe(true);
  });
});
