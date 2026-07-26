import { describe, expect, it } from "vitest";

import type { GroupElement, Transform } from "../model/shape.js";
import type { Slide } from "../model/slide.js";
import { asEmu } from "../utils/unit-types.js";
import { renderSlideToSvg } from "./svg-renderer.js";

const EMU_PER_PIXEL = 9525;

function transform(
  x: number,
  y: number,
  width: number,
  height: number,
  options: Partial<Pick<Transform, "rotation" | "flipH" | "flipV">> = {},
): Transform {
  return {
    offsetX: asEmu(x * EMU_PER_PIXEL),
    offsetY: asEmu(y * EMU_PER_PIXEL),
    extentWidth: asEmu(width * EMU_PER_PIXEL),
    extentHeight: asEmu(height * EMU_PER_PIXEL),
    rotation: 0,
    flipH: false,
    flipV: false,
    ...options,
  };
}

function group(overrides: Partial<GroupElement>): GroupElement {
  return {
    type: "group",
    transform: transform(0, 0, 1, 1),
    childTransform: transform(0, 0, 1, 1),
    children: [],
    effects: null,
    ...overrides,
  };
}

describe("renderSlideToSvg group transforms", () => {
  it("maps non-matching chOff/chExt with scale followed by child-origin translation", () => {
    const scaled = group({
      transform: transform(10, 20, 200, 120),
      childTransform: transform(5, 7, 100, 40),
    });

    const svg = renderSlideToSvg(
      { slideNumber: 1, background: null, elements: [scaled], showMasterSp: true },
      { width: asEmu(400 * EMU_PER_PIXEL), height: asEmu(300 * EMU_PER_PIXEL) },
    );

    expect(svg).toContain('<g transform="translate(10, 20) scale(2, 3) translate(-5, -7)">');
  });

  it("composes nested rotation, flip, and non-uniform scale from outer to inner groups", () => {
    const inner = group({
      transform: transform(20, 10, 80, 120, { rotation: -30, flipV: true }),
      childTransform: transform(2, 3, 40, 30),
    });
    const outer = group({
      transform: transform(10, 20, 200, 100, { rotation: 90, flipH: true }),
      childTransform: transform(5, 7, 100, 50),
      children: [inner],
    });
    const slide: Slide = {
      slideNumber: 1,
      background: null,
      elements: [outer],
      showMasterSp: true,
    };

    const svg = renderSlideToSvg(slide, {
      width: asEmu(500 * EMU_PER_PIXEL),
      height: asEmu(400 * EMU_PER_PIXEL),
    });

    const outerTransform =
      "translate(10, 20) rotate(90, 100, 50) translate(200, 0) scale(-1, 1) scale(2, 2) translate(-5, -7)";
    const innerTransform =
      "translate(20, 10) rotate(-30, 40, 60) translate(0, 120) scale(1, -1) scale(2, 4) translate(-2, -3)";
    expect(svg).toContain(
      `<g transform="${outerTransform}"><g transform="${innerTransform}"></g></g>`,
    );
  });

  it("uses identity scale only for a zero child extent axis", () => {
    const zeroWidth = group({
      transform: transform(10, 20, 200, 120),
      childTransform: transform(5, 7, 0, 40),
    });

    const svg = renderSlideToSvg(
      { slideNumber: 1, background: null, elements: [zeroWidth], showMasterSp: true },
      { width: asEmu(400 * EMU_PER_PIXEL), height: asEmu(300 * EMU_PER_PIXEL) },
    );

    expect(svg).toContain('<g transform="translate(10, 20) scale(1, 3) translate(-5, -7)">');
  });
});
