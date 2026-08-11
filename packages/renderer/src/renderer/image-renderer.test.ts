import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  createRepresentativeEmf,
  createRepresentativeWmf,
} from "../../../../vrt/snapshot/fixtures-src/images.js";
import type { ImageElement } from "../model/image.js";
import type { Transform } from "../model/shape.js";
import { unsafeFixtureAssertion } from "../unsafe-type-assertion.js";
import { createWarningLogger } from "../warning-logger.js";
import { renderImage } from "./image-renderer.js";
import { createRendererContext } from "./render-context.js";

beforeEach(() => {
  let counter = 0;
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    return unsafeFixtureAssertion<ReturnType<typeof crypto.randomUUID>>(`test-uuid-${counter++}`);
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

describe("EMF rendering", () => {
  it("converts EMF content before applying the existing crop contract", () => {
    const warnings = createWarningLogger("warn");
    const result = renderImage(
      makeImage({
        mimeType: "image/emf",
        imageData: createRepresentativeEmf().toString("base64"),
        srcRect: { left: 0.1, top: 0, right: 0, bottom: 0 },
      }),
      createRendererContext({ warningLogger: warnings }),
    );

    expect(result.content).toContain('viewBox="0 0 1000 1000"');
    expect(result.content).toContain(">EMF</text>");
    expect(result.content).toContain("clip-path=");
    expect(result.content).not.toContain("[EMF]");
    expect(warnings.getWarningEntries()).toEqual([]);
  });

  it("renders valid WMF content instead of a placeholder", () => {
    const result = renderImage(
      makeImage({
        mimeType: "image/wmf",
        imageData: createRepresentativeWmf().toString("base64"),
      }),
    );

    expect(result.content).toContain('viewBox="0 0 1000 1000"');
    expect(result.content).toContain(">WMF</text>");
    expect(result.content).not.toContain("[WMF]");
  });

  it("applies stretch insets to converted metafile SVG", () => {
    const result = renderImage(
      makeImage({
        mimeType: "image/emf",
        imageData: createRepresentativeEmf().toString("base64"),
        stretch: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
      }),
    );

    expect(result.content).toContain('<svg viewBox="0 0 1000 1000"');
    expect(result.content).toContain('x="19" y="29" width="115" height="101"');
  });

  it.each([
    ["x", "translate(96, 0) scale(-1, 1)"],
    ["y", "translate(0, 72) scale(1, -1)"],
    ["xy", "translate(96, 72) scale(-1, -1)"],
  ] as const)("applies tile flip %s to converted WMF SVG", (flip, transform) => {
    const result = renderImage(
      makeImage({
        mimeType: "image/wmf",
        imageData: createRepresentativeWmf().toString("base64"),
        tile: { tx: 0, ty: 0, sx: 0.5, sy: 0.5, flip, align: "tl" },
      }),
    );

    expect(result.defs[0]).toContain(`transform="${transform}"`);
    expect(result.defs[0]).toContain(">WMF</text>");
  });

  it("applies shape and blip effects around converted EMF SVG", () => {
    const result = renderImage(
      makeImage({
        mimeType: "image/emf",
        imageData: createRepresentativeEmf().toString("base64"),
        effects: {
          outerShadow: null,
          innerShadow: null,
          glow: { radius: 50800, color: { hex: "#FF0000", alpha: 0.5 } },
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
    expect(result.content).toContain('filter="url(#effect-');
    expect(result.content).toContain('filter="url(#blip-effect-');
    expect(result.content).toContain(">EMF</text>");
  });

  it("caches a successful conversion for repeated images in one render context", () => {
    const context = createRendererContext();
    const image = makeImage({
      mimeType: "image/emf",
      imageData: createRepresentativeEmf().toString("base64"),
    });

    renderImage(image, context);
    renderImage(image, context);

    expect(context.metafileConversionCache.size).toBe(1);
  });

  it("caches a failed conversion while warning for each repeated fallback", () => {
    const warnings = createWarningLogger("warn");
    const context = createRendererContext({ warningLogger: warnings });
    const image = makeImage({ mimeType: "image/emf", imageData: "broken" });

    renderImage(image, context);
    renderImage(image, context);

    expect(context.metafileConversionCache.size).toBe(1);
    expect(
      warnings.getWarningEntries().filter((entry) => entry.feature === "image.metafile-conversion"),
    ).toHaveLength(2);
  });
});

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
    const warnings = createWarningLogger("warn");
    const result = renderImage(
      makeImage({ mimeType: "image/wmf" }),
      createRendererContext({ warningLogger: warnings }),
    );
    expect(result.content).toContain("[WMF]");
    expect(result.content).not.toContain("<image");
    const warning = warnings
      .getWarningEntries()
      .find((entry) => entry.feature === "image.metafile-conversion");
    expect(warning?.message).toContain("WMF conversion invalid-data");
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
