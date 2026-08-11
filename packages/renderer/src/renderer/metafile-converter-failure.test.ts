import { describe, expect, it, vi } from "vitest";

import { createRepresentativeEmf } from "../../../../vrt/snapshot/fixtures-src/images.js";

const throwingRenderer = vi.hoisted(() => ({
  loggingEnabled: vi.fn(),
  Renderer: class {
    render(): never {
      throw new Error("synthetic parser failure");
    }
  },
}));

vi.mock("rtf.js/dist/EMFJS.bundle.js", () => throwingRenderer);
vi.mock("rtf.js/dist/WMFJS.bundle.js", () => throwingRenderer);

import { convertMetafileToSvgData } from "./metafile-converter.js";

describe("metafile parser failure", () => {
  it("returns conversion-failed without throwing", () => {
    expect(
      convertMetafileToSvgData(createRepresentativeEmf().toString("base64"), "image/emf"),
    ).toMatchObject({
      ok: false,
      reason: "conversion-failed",
      message: "synthetic parser failure",
    });
  });
});
