import { describe, expect, it, vi } from "vitest";

import { createRepresentativeEmf } from "../../../../vrt/snapshot/fixtures-src/images.js";
import type { ImageElement } from "../model/image.js";
import { asEmu } from "../utils/unit-types.js";

const rendererWithReferences = vi.hoisted(() => ({
  loggingEnabled: vi.fn(),
  Renderer: class {
    render(): SVGSVGElement {
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("viewBox", "0 0 1000 1000");
      const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");
      const gradient = document.createElementNS("http://www.w3.org/2000/svg", "linearGradient");
      gradient.setAttribute("id", "paint");
      defs.appendChild(gradient);
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("fill", "url(#paint)");
      svg.appendChild(defs);
      svg.appendChild(rect);
      return svg;
    }
  },
}));

vi.mock("rtf.js/dist/EMFJS.bundle.js", () => rendererWithReferences);
vi.mock("rtf.js/dist/WMFJS.bundle.js", () => rendererWithReferences);

import { renderImage } from "./image-renderer.js";
import { createRendererContext } from "./render-context.js";

describe("cached metafile SVG insertion", () => {
  it("reuses conversion while assigning distinct deterministic fragment namespaces", () => {
    const context = createRendererContext();
    const image: ImageElement = {
      type: "image",
      transform: {
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        extentWidth: asEmu(914400),
        extentHeight: asEmu(914400),
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      imageData: createRepresentativeEmf().toString("base64"),
      mimeType: "image/emf",
      effects: null,
      blipEffects: null,
      srcRect: null,
      stretch: null,
      tile: null,
    };

    const first = renderImage(image, context).content;
    const second = renderImage(image, context).content;

    expect(context.metafileConversionCache.size).toBe(1);
    expect(first).toContain('id="metafile-0-paint"');
    expect(first).toContain('fill="url(#metafile-0-paint)"');
    expect(second).toContain('id="metafile-1-paint"');
    expect(second).toContain('fill="url(#metafile-1-paint)"');
  });
});
