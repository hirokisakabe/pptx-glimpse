import { describe, expect, it } from "vitest";

import {
  createRepresentativeEmf,
  createRepresentativeWmf,
} from "../../../../vrt/snapshot/fixtures-src/images.js";
import { convertMetafileToSvgData } from "./metafile-converter.js";

describe("convertMetafileToSvgData", () => {
  it.each([
    ["EMF", createRepresentativeEmf, "image/emf"],
    ["WMF", createRepresentativeWmf, "image/wmf"],
  ] as const)("converts a representative %s stream to SVG", (label, createFixture, mimeType) => {
    const result = convertMetafileToSvgData(createFixture().toString("base64"), mimeType);

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const svg = Buffer.from(result.imageData, "base64").toString("utf8");
    expect(result.mimeType).toBe("image/svg+xml");
    expect(svg).toContain("<svg");
    expect(svg).toMatch(/<(?:path|rect|polygon)\b/);
    expect(svg).toContain(label);
  });

  it("identifies structurally unknown records without throwing", () => {
    const bytes = createRepresentativeEmf();
    bytes.writeUInt32LE(0x7fff, 88);

    expect(convertMetafileToSvgData(bytes.toString("base64"), "image/emf")).toMatchObject({
      ok: false,
      reason: "unsupported-record",
    });
  });

  it("identifies malformed data without throwing", () => {
    expect(
      convertMetafileToSvgData(Buffer.from("broken").toString("base64"), "image/wmf"),
    ).toMatchObject({ ok: false, reason: "invalid-data" });
  });
});
