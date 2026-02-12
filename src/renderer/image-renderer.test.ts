import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderImage } from "./image-renderer.js";
import type { ImageElement } from "../model/image.js";
import type { Transform } from "../model/shape.js";

beforeEach(() => {
  let counter = 0;
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    return `test-uuid-${counter++}` as ReturnType<typeof crypto.randomUUID>;
  });
});

function makeTransform(overrides: Partial<Transform> = {}): Transform {
  return {
    offsetX: 914400,
    offsetY: 914400,
    extentWidth: 1828800,
    extentHeight: 1371600,
    rotation: 0,
    flipH: false,
    flipV: false,
    ...overrides,
  };
}

function makeImage(overrides: Partial<ImageElement> = {}): ImageElement {
  return {
    type: "image",
    transform: makeTransform(),
    imageData: "iVBORw0KGgo=",
    mimeType: "image/png",
    effects: null,
    ...overrides,
  };
}

describe("renderImage", () => {
  it("renders basic image element", () => {
    const result = renderImage(makeImage());

    expect(result).toContain('<g transform="translate(96, 96)">');
    expect(result).toContain('href="data:image/png;base64,iVBORw0KGgo="');
    expect(result).toContain('width="192"');
    expect(result).toContain('height="144"');
    expect(result).toContain('preserveAspectRatio="none"');
    expect(result).toContain("</g>");
  });

  it("renders image with JPEG mime type", () => {
    const result = renderImage(makeImage({ mimeType: "image/jpeg", imageData: "/9j/4AAQ=" }));
    expect(result).toContain('href="data:image/jpeg;base64,/9j/4AAQ="');
  });

  it("renders image with rotation", () => {
    const result = renderImage(makeImage({ transform: makeTransform({ rotation: 45 }) }));
    expect(result).toContain("rotate(45, 96, 72)");
  });

  it("renders image with flipH", () => {
    const result = renderImage(makeImage({ transform: makeTransform({ flipH: true }) }));
    expect(result).toContain("translate(192, 0) scale(-1, 1)");
  });

  it("does not include filter when effects is null", () => {
    const result = renderImage(makeImage());
    expect(result).not.toContain("<filter");
    expect(result).not.toContain("filter=");
  });

  it("renders image with effects", () => {
    const result = renderImage(
      makeImage({
        effects: {
          outerShadow: {
            blurRadius: 50800,
            distance: 38100,
            direction: 45,
            color: { hex: "#000000", alpha: 0.5 },
            alignment: "br",
            rotateWithShape: false,
          },
          innerShadow: null,
          glow: null,
          softEdge: null,
        },
      }),
    );

    expect(result).toContain('<filter id="effect-test-uuid-0"');
    expect(result).toContain('filter="url(#effect-test-uuid-0)"');
  });
});
