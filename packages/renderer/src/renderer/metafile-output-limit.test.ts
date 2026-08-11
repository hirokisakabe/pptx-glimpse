import { describe, expect, it, vi } from "vitest";

import { createRepresentativeEmf } from "../../../../vrt/snapshot/fixtures-src/images.js";

const oversizedRenderer = vi.hoisted(() => ({
  mode: "output",
  loggingEnabled: vi.fn(),
  Renderer: class {
    render(): SVGSVGElement {
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("viewBox", "0 0 1000 1000");
      if (oversizedRenderer.mode === "output") {
        svg.textContent = "x".repeat(8 * 1024 * 1024 + 1);
      } else {
        for (let index = 0; index < 100_000; index++) {
          svg.appendChild(document.createElementNS("http://www.w3.org/2000/svg", "path"));
        }
      }
      return svg;
    }
  },
}));

vi.mock("rtf.js/dist/EMFJS.bundle.js", () => oversizedRenderer);
vi.mock("rtf.js/dist/WMFJS.bundle.js", () => oversizedRenderer);

import { convertMetafileToSvgData } from "./metafile-converter.js";

describe("metafile SVG output limits", () => {
  it("fails conversion with a stable diagnostic when rtf.js produces oversized SVG", () => {
    oversizedRenderer.mode = "output";
    expect(
      convertMetafileToSvgData(createRepresentativeEmf().toString("base64"), "image/emf"),
    ).toMatchObject({
      ok: false,
      reason: "conversion-failed",
      message: "Converted metafile SVG output limit exceeded.",
    });
  });

  it("fails conversion before serialization when rtf.js produces too many SVG nodes", () => {
    oversizedRenderer.mode = "nodes";

    expect(
      convertMetafileToSvgData(createRepresentativeEmf().toString("base64"), "image/emf"),
    ).toMatchObject({
      ok: false,
      reason: "conversion-failed",
      message: "Converted metafile SVG node limit exceeded.",
    });
  });
});
