import { describe, expect, it } from "vitest";

import {
  createRepresentativeEmf,
  createRepresentativeWmf,
} from "../../../../vrt/snapshot/fixtures-src/images.js";
import { convertMetafileToSvgData, inlineSvgData } from "./metafile-converter.js";

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

  it("uses non-1000 EMF header bounds as the SVG viewBox", () => {
    const bytes = createRepresentativeEmf();
    writeInt32Rectangle(bytes, 8, 120, 40, 1720, 740);

    const result = convertMetafileToSvgData(bytes.toString("base64"), "image/emf");

    expect(decodedSvg(result)).toContain('viewBox="120 40 1600 700"');
  });

  it("maps EMF text origin, font size, rotation, weight, and offDx advances", () => {
    const result = convertMetafileToSvgData(
      createRepresentativeEmf().toString("base64"),
      "image/emf",
    );

    const svg = decodedSvg(result);
    expect(svg).toContain('<text x="250" y="820"');
    expect(svg).toContain('dx="0 150 150"');
    expect(svg).toContain('font-family="Noto Sans"');
    expect(svg).toContain('font-size="72"');
    expect(svg).toContain('font-weight="bold"');
    expect(svg).toContain('transform="rotate(-12 250 820)"');
  });

  it.each([
    ["non-compatible graphics mode", 0x54, 24, 2],
    ["explicit font width", 0x52, 16, 20],
  ] as const)("rejects EMF text with unsupported %s", (_label, recordType, fieldOffset, value) => {
    const bytes = createRepresentativeEmf();
    bytes.writeInt32LE(value, findEmfRecord(bytes, recordType) + fieldOffset);

    expect(convertMetafileToSvgData(bytes.toString("base64"), "image/emf")).toMatchObject({
      ok: false,
      reason: "unsupported-record",
    });
  });

  it("rejects an out-of-range EMF offDx array", () => {
    const bytes = createRepresentativeEmf();
    bytes.writeUInt32LE(0xfffffff0, findEmfRecord(bytes, 0x54) + 72);

    expect(convertMetafileToSvgData(bytes.toString("base64"), "image/emf")).toMatchObject({
      ok: false,
      reason: "invalid-data",
      message: expect.stringContaining("advance range"),
    });
  });

  it("uses non-1000 WMF stream window extents as the SVG viewBox", () => {
    const bytes = createRepresentativeWmf();
    // META_SETWINDOWEXT follows the 18-byte header and 10-byte META_SETWINDOWORG record.
    bytes.writeInt16LE(700, 34);
    bytes.writeInt16LE(1600, 36);

    const result = convertMetafileToSvgData(bytes.toString("base64"), "image/wmf");

    expect(decodedSvg(result)).toContain('viewBox="0 0 1600 700"');
  });

  it("uses a nonzero placeable WMF bounding box as the SVG viewBox", () => {
    const bytes = withPlaceableHeader(createRepresentativeWmf(), 50, 75, 1550, 675);

    const result = convertMetafileToSvgData(bytes.toString("base64"), "image/wmf");

    expect(decodedSvg(result)).toContain('viewBox="50 75 1500 600"');
  });

  it.each([
    ["EMF", "image/emf", emfWithTrailingBytes],
    ["WMF", "image/wmf", wmfWithTrailingBytes],
  ] as const)("rejects trailing bytes after the %s EOF record", (_label, mimeType, fixture) => {
    expect(convertMetafileToSvgData(fixture().toString("base64"), mimeType)).toMatchObject({
      ok: false,
      reason: "invalid-data",
      message: expect.stringContaining("trailing bytes"),
    });
  });

  it.each([
    ["EMF", "image/emf", emfWithoutEof],
    ["WMF", "image/wmf", wmfWithoutEof],
  ] as const)("requires an explicit %s EOF record", (_label, mimeType, fixture) => {
    expect(convertMetafileToSvgData(fixture().toString("base64"), mimeType)).toMatchObject({
      ok: false,
      reason: "invalid-data",
      message: expect.stringContaining("EOF record is missing"),
    });
  });

  it("validates the EMF declared byte size and record count", () => {
    const wrongSize = createRepresentativeEmf();
    wrongSize.writeUInt32LE(wrongSize.length - 4, 48);
    const wrongCount = createRepresentativeEmf();
    wrongCount.writeUInt32LE(wrongCount.readUInt32LE(52) + 1, 52);

    expect(convertMetafileToSvgData(wrongSize.toString("base64"), "image/emf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("declared byte size"),
    });
    expect(convertMetafileToSvgData(wrongCount.toString("base64"), "image/emf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("declared record count"),
    });
  });

  it("validates WMF declared byte and maximum-record sizes", () => {
    const wrongSize = createRepresentativeWmf();
    wrongSize.writeUInt32LE(wrongSize.readUInt32LE(6) - 1, 6);
    const wrongMaximum = createRepresentativeWmf();
    wrongMaximum.writeUInt32LE(wrongMaximum.readUInt32LE(12) + 1, 12);

    expect(convertMetafileToSvgData(wrongSize.toString("base64"), "image/wmf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("declared byte size"),
    });
    expect(convertMetafileToSvgData(wrongMaximum.toString("base64"), "image/wmf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("maximum record size"),
    });
  });

  it("validates EOF-declared sizes and the placeable WMF checksum", () => {
    const emf = createRepresentativeEmf();
    const emfEof = findEmfRecord(emf, 0x0e);
    emf.writeUInt32LE(emf.readUInt32LE(emfEof + 4) + 4, emf.length - 4);

    const wmf = createRepresentativeWmf();
    wmf.writeUInt32LE(4, wmf.length - 6);

    const placeable = withPlaceableHeader(createRepresentativeWmf(), 0, 0, 1000, 1000);
    placeable.writeUInt16LE(placeable.readUInt16LE(20) ^ 1, 20);

    expect(convertMetafileToSvgData(emf.toString("base64"), "image/emf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("EOF record has an invalid declared size"),
    });
    expect(convertMetafileToSvgData(wmf.toString("base64"), "image/wmf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("EOF record has an invalid declared size"),
    });
    expect(convertMetafileToSvgData(placeable.toString("base64"), "image/wmf")).toMatchObject({
      ok: false,
      message: expect.stringContaining("checksum"),
    });
  });

  it("escapes allowlisted inline SVG attributes and rejects all other names", () => {
    const encoded = Buffer.from('<svg viewBox="0 0 1 1"></svg>').toString("base64");

    expect(inlineSvgData(encoded, { x: '1"<&', width: 2 })).toContain(
      'x="1&quot;&lt;&amp;" width="2"',
    );
    expect(() => inlineSvgData(encoded, { onload: "alert(1)" })).toThrow(
      "Unsupported inline SVG attribute 'onload'.",
    );
  });
});

function decodedSvg(result: ReturnType<typeof convertMetafileToSvgData>): string {
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error(result.message);
  return Buffer.from(result.imageData, "base64").toString("utf8");
}

function findEmfRecord(buffer: Buffer, recordType: number): number {
  let offset = 0;
  while (offset + 8 <= buffer.length) {
    if (buffer.readUInt32LE(offset) === recordType) return offset;
    offset += buffer.readUInt32LE(offset + 4);
  }
  throw new Error(`EMF record 0x${recordType.toString(16)} not found`);
}

function writeInt32Rectangle(
  buffer: Buffer,
  offset: number,
  left: number,
  top: number,
  right: number,
  bottom: number,
): void {
  [left, top, right, bottom].forEach((value, index) =>
    buffer.writeInt32LE(value, offset + index * 4),
  );
}

function withPlaceableHeader(
  metafile: Buffer,
  left: number,
  top: number,
  right: number,
  bottom: number,
): Buffer {
  const header = Buffer.alloc(22);
  header.writeUInt32LE(0x9ac6cdd7, 0);
  header.writeUInt16LE(0, 4);
  header.writeInt16LE(left, 6);
  header.writeInt16LE(top, 8);
  header.writeInt16LE(right, 10);
  header.writeInt16LE(bottom, 12);
  header.writeUInt16LE(1440, 14);
  header.writeUInt32LE(0, 16);
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= header.readUInt16LE(offset);
  header.writeUInt16LE(checksum, 20);
  return Buffer.concat([header, metafile]);
}

function emfWithTrailingBytes(): Buffer {
  const bytes = Buffer.concat([createRepresentativeEmf(), Buffer.alloc(4)]);
  bytes.writeUInt32LE(bytes.length, 48);
  return bytes;
}

function wmfWithTrailingBytes(): Buffer {
  const bytes = Buffer.concat([createRepresentativeWmf(), Buffer.alloc(2)]);
  bytes.writeUInt32LE(bytes.length / 2, 6);
  return bytes;
}

function emfWithoutEof(): Buffer {
  const original = createRepresentativeEmf();
  const bytes = original.subarray(0, original.length - 20);
  bytes.writeUInt32LE(bytes.length, 48);
  bytes.writeUInt32LE(bytes.readUInt32LE(52) - 1, 52);
  return bytes;
}

function wmfWithoutEof(): Buffer {
  const original = createRepresentativeWmf();
  const bytes = original.subarray(0, original.length - 6);
  bytes.writeUInt32LE(bytes.length / 2, 6);
  return bytes;
}
