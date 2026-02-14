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
    blipEffects: null,
    srcRect: null,
    stretch: null,
    tile: null,
    ...overrides,
  };
}

describe("renderImage", () => {
  it("renders basic image element", () => {
    const result = renderImage(makeImage());

    expect(result.content).toContain('<g transform="translate(96, 96)">');
    expect(result.content).toContain('href="data:image/png;base64,iVBORw0KGgo="');
    expect(result.content).toContain('width="192"');
    expect(result.content).toContain('height="144"');
    expect(result.content).toContain('preserveAspectRatio="none"');
    expect(result.content).toContain("</g>");
    expect(result.defs).toHaveLength(0);
  });

  it("renders image with JPEG mime type", () => {
    const result = renderImage(makeImage({ mimeType: "image/jpeg", imageData: "/9j/4AAQ=" }));
    expect(result.content).toContain('href="data:image/jpeg;base64,/9j/4AAQ="');
  });

  it("renders image with rotation", () => {
    const result = renderImage(makeImage({ transform: makeTransform({ rotation: 45 }) }));
    expect(result.content).toContain("rotate(45, 96, 72)");
  });

  it("renders image with flipH", () => {
    const result = renderImage(makeImage({ transform: makeTransform({ flipH: true }) }));
    expect(result.content).toContain("translate(192, 0) scale(-1, 1)");
  });

  it("does not include filter when effects is null", () => {
    const result = renderImage(makeImage());
    expect(result.content).not.toContain("<filter");
    expect(result.content).not.toContain("filter=");
    expect(result.defs).toHaveLength(0);
  });

  it("renders image with srcRect crop", () => {
    const result = renderImage(
      makeImage({ srcRect: { left: 0.1, top: 0.2, right: 0.1, bottom: 0.2 } }),
    );

    // clipPath def should be present
    expect(result.content).toContain("<clipPath");
    expect(result.content).toContain('width="192"');
    expect(result.content).toContain('height="144"');

    // image should be scaled up: 192 / (1 - 0.1 - 0.1) = 240, 144 / (1 - 0.2 - 0.2) = 240
    expect(result.content).toContain('width="240"');
    expect(result.content).toContain('height="240"');

    // image offset: x = -0.1 * 240 = -24, y = -0.2 * 240 = -48
    expect(result.content).toContain('x="-24"');
    expect(result.content).toContain('y="-48"');

    expect(result.content).toContain("clip-path=");
  });

  it("does not include clipPath when srcRect is null", () => {
    const result = renderImage(makeImage());
    expect(result.content).not.toContain("<clipPath");
    expect(result.content).not.toContain("clip-path=");
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

    expect(result.defs).toHaveLength(1);
    expect(result.defs[0]).toContain('<filter id="effect-test-uuid-0"');
    expect(result.content).toContain('filter="url(#effect-test-uuid-0)"');
  });

  it("renders EMF placeholder", () => {
    const result = renderImage(makeImage({ mimeType: "image/emf" }));
    expect(result.content).toContain('fill="#E0E0E0"');
    expect(result.content).toContain("[EMF]");
    expect(result.content).toContain("<rect");
    expect(result.content).toContain("<text");
    expect(result.content).not.toContain("<image");
    expect(result.defs).toHaveLength(0);
  });

  it("renders WMF placeholder", () => {
    const result = renderImage(makeImage({ mimeType: "image/wmf" }));
    expect(result.content).toContain("[WMF]");
    expect(result.content).not.toContain("<image");
  });

  it("renders image with blipEffects grayscale", () => {
    const result = renderImage(
      makeImage({
        blipEffects: {
          grayscale: true,
          biLevel: null,
          blur: null,
          lum: null,
          duotone: null,
        },
      }),
    );
    expect(result.defs).toHaveLength(1);
    expect(result.defs[0]).toContain('<filter id="blip-effect-');
    expect(result.defs[0]).toContain('type="saturate" values="0"');
    expect(result.content).toContain('filter="url(#blip-effect-');
  });

  it("renders image with stretch fillRect", () => {
    const result = renderImage(
      makeImage({
        stretch: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 },
      }),
    );
    // 192 * 0.1 = 19, 144 * 0.1 = 14
    expect(result.content).toContain('x="19"');
    expect(result.content).toContain('y="14"');
    // 192 * 0.8 = 154, 144 * 0.8 = 115
    expect(result.content).toContain('width="154"');
    expect(result.content).toContain('height="115"');
  });

  it("renders image with tile", () => {
    const result = renderImage(
      makeImage({
        tile: { tx: 0, ty: 0, sx: 0.5, sy: 0.5, flip: "none", align: "tl" },
      }),
    );
    expect(result.defs).toHaveLength(1);
    expect(result.defs[0]).toContain("<pattern");
    expect(result.defs[0]).toContain("patternUnits=");
    expect(result.defs[0]).toContain('width="96"');
    expect(result.defs[0]).toContain('height="72"');
    expect(result.content).toContain("<rect");
    expect(result.content).toContain('fill="url(#tile-');
  });

  it("renders tiled image with flip x", () => {
    const result = renderImage(
      makeImage({
        tile: { tx: 0, ty: 0, sx: 0.5, sy: 0.5, flip: "x", align: "tl" },
      }),
    );
    expect(result.defs[0]).toContain("scale(-1, 1)");
  });

  it("applies both blipEffects and effectLst with nested g", () => {
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
        blipEffects: {
          grayscale: true,
          biLevel: null,
          blur: null,
          lum: null,
          duotone: null,
        },
      }),
    );
    expect(result.defs).toHaveLength(2);
    expect(result.defs.some((d) => d.includes('<filter id="effect-'))).toBe(true);
    expect(result.defs.some((d) => d.includes('<filter id="blip-effect-'))).toBe(true);
    expect(result.content).toContain('filter="url(#effect-');
    expect(result.content).toContain('filter="url(#blip-effect-');
  });
});
