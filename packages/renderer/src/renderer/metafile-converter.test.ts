import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";

import {
  createRepresentativeEmf,
  createRepresentativeWmf,
} from "../../../../vrt/snapshot/fixtures-src/images.js";
import { createWarningLogger } from "../warning-logger.js";
import {
  BoundedMetafileConversionCache,
  convertMetafileToSvgData,
  inlineSvgData,
  resolveMetafileImageSource,
} from "./metafile-converter.js";

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

    expectFailureMessage(
      convertMetafileToSvgData(bytes.toString("base64"), "image/emf"),
      "advance range",
    );
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
    expectFailureMessage(
      convertMetafileToSvgData(fixture().toString("base64"), mimeType),
      "trailing bytes",
    );
  });

  it.each([
    ["EMF", "image/emf", emfWithoutEof],
    ["WMF", "image/wmf", wmfWithoutEof],
  ] as const)("requires an explicit %s EOF record", (_label, mimeType, fixture) => {
    expectFailureMessage(
      convertMetafileToSvgData(fixture().toString("base64"), mimeType),
      "EOF record is missing",
    );
  });

  it("validates the EMF declared byte size and record count", () => {
    const wrongSize = createRepresentativeEmf();
    wrongSize.writeUInt32LE(wrongSize.length - 4, 48);
    const wrongCount = createRepresentativeEmf();
    wrongCount.writeUInt32LE(wrongCount.readUInt32LE(52) + 1, 52);

    expectFailureMessage(
      convertMetafileToSvgData(wrongSize.toString("base64"), "image/emf"),
      "declared byte size",
    );
    expectFailureMessage(
      convertMetafileToSvgData(wrongCount.toString("base64"), "image/emf"),
      "declared record count",
    );
  });

  it("validates WMF declared byte and maximum-record sizes", () => {
    const wrongSize = createRepresentativeWmf();
    wrongSize.writeUInt32LE(wrongSize.readUInt32LE(6) - 1, 6);
    const wrongMaximum = createRepresentativeWmf();
    wrongMaximum.writeUInt32LE(wrongMaximum.readUInt32LE(12) + 1, 12);

    expectFailureMessage(
      convertMetafileToSvgData(wrongSize.toString("base64"), "image/wmf"),
      "declared byte size",
    );
    expectFailureMessage(
      convertMetafileToSvgData(wrongMaximum.toString("base64"), "image/wmf"),
      "maximum record size",
    );
  });

  it("validates EOF-declared sizes and the placeable WMF checksum", () => {
    const emf = createRepresentativeEmf();
    const emfEof = findEmfRecord(emf, 0x0e);
    emf.writeUInt32LE(emf.readUInt32LE(emfEof + 4) + 4, emf.length - 4);

    const wmf = createRepresentativeWmf();
    wmf.writeUInt32LE(4, wmf.length - 6);

    const placeable = withPlaceableHeader(createRepresentativeWmf(), 0, 0, 1000, 1000);
    placeable.writeUInt16LE(placeable.readUInt16LE(20) ^ 1, 20);

    expectFailureMessage(
      convertMetafileToSvgData(emf.toString("base64"), "image/emf"),
      "EOF record has an invalid declared size",
    );
    expectFailureMessage(
      convertMetafileToSvgData(wmf.toString("base64"), "image/wmf"),
      "EOF record has an invalid declared size",
    );
    expectFailureMessage(
      convertMetafileToSvgData(placeable.toString("base64"), "image/wmf"),
      "checksum",
    );
  });

  it("reads EMR_EOF SizeLast from the last DWORD and validates optional palette data", () => {
    const valid = emfWithPaletteEof();
    const malformedOffset = Buffer.from(valid);
    const malformedCount = Buffer.from(valid);
    const malformedSize = Buffer.from(valid);
    const eof = findEmfRecord(valid, 0x0e);
    malformedOffset.writeUInt32LE(24, eof + 12);
    malformedCount.writeUInt32LE(3, eof + 8);
    malformedSize.writeUInt32LE(20, eof + valid.readUInt32LE(eof + 4) - 4);

    expect(convertMetafileToSvgData(valid.toString("base64"), "image/emf")).toMatchObject({
      ok: true,
    });
    expectFailureMessage(
      convertMetafileToSvgData(malformedOffset.toString("base64"), "image/emf"),
      "palette declaration",
    );
    expectFailureMessage(
      convertMetafileToSvgData(malformedCount.toString("base64"), "image/emf"),
      "palette declaration",
    );
    expectFailureMessage(
      convertMetafileToSvgData(malformedSize.toString("base64"), "image/emf"),
      "EOF record has an invalid declared size",
    );
  });

  it.each([
    ["top-relative", -1, "#00FF00"],
    ["multi-relative", -2, "#333333"],
  ] as const)(
    "restores mapping, selected font, text color, and alignment with %s RestoreDC",
    (_label, restoreValue, restoredColor) => {
      const bytes = emfWithRestoredTextState(restoreValue);

      const svg = decodedSvg(convertMetafileToSvgData(bytes.toString("base64"), "image/emf"));

      expect(svg).toContain('<text x="250" y="820"');
      expect(svg).toContain(`fill="${restoredColor}"`);
      expect(svg).toContain('font-family="Noto Sans"');
      expect(svg).not.toContain("Changed Font");
    },
  );

  it.each([0, 1])("rejects non-negative RestoreDC SavedDC value %i", (restoreValue) => {
    expectFailureMessage(
      convertMetafileToSvgData(
        emfWithRestoredTextState(restoreValue).toString("base64"),
        "image/emf",
      ),
      "SavedDC must be negative",
    );
  });

  it("rejects RestoreDC references outside the saved-state stack", () => {
    const bytes = createRepresentativeEmf();
    const textOffset = findEmfRecord(bytes, 0x54);
    const restore = emfRecord(0x22, int32Values(-1));
    const malformed = Buffer.concat([
      bytes.subarray(0, textOffset),
      restore,
      bytes.subarray(textOffset),
    ]);
    malformed.writeUInt32LE(malformed.length, 48);
    malformed.writeUInt32LE(malformed.readUInt32LE(52) + 1, 52);

    expectFailureMessage(
      convertMetafileToSvgData(malformed.toString("base64"), "image/emf"),
      "invalid saved state",
    );
  });

  it("sanitizes XML 1.0 forbidden controls in EMF text and font names", () => {
    const bytes = createRepresentativeEmf();
    const font = findEmfRecord(bytes, 0x52);
    const text = findEmfRecord(bytes, 0x54);
    bytes.writeUInt16LE(1, font + 40);
    bytes.writeUInt16LE(1, text + 76);

    const svg = decodedSvg(convertMetafileToSvgData(bytes.toString("base64"), "image/emf"));

    expect(svg).not.toMatch(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/);
    expect(svg).toContain("�");
    expect(new DOMParser().parseFromString(svg, "image/svg+xml").documentElement.tagName).toBe(
      "svg",
    );
  });

  it("rejects oversized encoded input and adversarial allocation declarations", () => {
    const geometry = createRepresentativeEmf();
    geometry.writeUInt32LE(0x02, findEmfRecord(geometry, 0x26));

    expectFailureMessage(
      convertMetafileToSvgData("A".repeat(12 * 1024 * 1024 + 1), "image/emf"),
      "encoded input limit",
    );
    expectFailureMessage(
      convertMetafileToSvgData(geometry.toString("base64"), "image/emf"),
      "geometry point limit",
    );
    expectFailureMessage(
      convertMetafileToSvgData(
        largeEmfRecord(0x51, 2 * 1024 * 1024 + 4).toString("base64"),
        "image/emf",
      ),
      "bitmap payload limit",
    );
    expectFailureMessage(
      convertMetafileToSvgData(
        largeEmfRecord(0x02, 4 * 1024 * 1024 + 4).toString("base64"),
        "image/emf",
      ),
      "record payload limit",
    );
  });

  it("rejects oversized encoded input before consulting the conversion cache", () => {
    const cache = {
      size: 0,
      get: () => {
        throw new Error("oversized input reached cache key lookup");
      },
      set: () => {
        throw new Error("oversized input reached cache storage");
      },
    };
    const warnings = createWarningLogger("warn");

    expect(
      resolveMetafileImageSource("A".repeat(12 * 1024 * 1024 + 1), "image/emf", warnings, cache),
    ).toBeUndefined();
    expect(warnings.getWarningEntries()).toHaveLength(1);
    expect(warnings.getWarningEntries()[0]?.message).toContain("encoded input limit");
  });

  it("bounds conversion cache entries while retaining compact keys", () => {
    const cache = new BoundedMetafileConversionCache();
    for (let index = 0; index < 20; index++) {
      cache.set(`image/emf:4:${index}:${index}`, {
        ok: false,
        reason: "invalid-data",
        message: "test",
      });
    }

    expect(cache.size).toBe(16);
    expect(cache.get("image/emf:4:0:0")).toBeUndefined();
    expect(cache.get("image/emf:4:19:19")).toMatchObject({ ok: false });
  });

  it("bounds conversion cache result memory", () => {
    const cache = new BoundedMetafileConversionCache();
    cache.set("first", {
      ok: true,
      imageData: "A".repeat(9 * 1024 * 1024),
      mimeType: "image/svg+xml",
    });
    cache.set("second", {
      ok: true,
      imageData: "B".repeat(9 * 1024 * 1024),
      mimeType: "image/svg+xml",
    });

    expect(cache.size).toBe(1);
    expect(cache.get("first")).toBeUndefined();
    expect(cache.get("second")).toMatchObject({ ok: true });
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

  it("deterministically namespaces IDs and all local fragment references", () => {
    const encoded = Buffer.from(
      '<svg><defs><linearGradient id="paint"><stop/></linearGradient><path id="shape" fill="url(#paint)"/></defs><use href="#shape" xlink:href="#shape"/></svg>',
    ).toString("base64");

    const first = inlineSvgData(encoded, {}, "metafile-0-");
    const second = inlineSvgData(encoded, {}, "metafile-1-");

    expect(first).toContain('id="metafile-0-paint"');
    expect(first).toContain('fill="url(#metafile-0-paint)"');
    expect(first).toContain('href="#metafile-0-shape"');
    expect(second).toContain('id="metafile-1-paint"');
    expect(second).not.toContain('id="metafile-0-paint"');
  });
});

function decodedSvg(result: ReturnType<typeof convertMetafileToSvgData>): string {
  if (!result.ok) throw new Error(result.message);
  expect(result.ok).toBe(true);
  return Buffer.from(result.imageData, "base64").toString("utf8");
}

function expectFailureMessage(
  result: ReturnType<typeof convertMetafileToSvgData>,
  message: string,
): void {
  expect(result.ok).toBe(false);
  if (result.ok) throw new Error("Expected metafile conversion to fail.");
  expect(result.reason).toBe("invalid-data");
  expect(result.message).toContain(message);
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

function emfWithPaletteEof(): Buffer {
  const original = createRepresentativeEmf();
  const eofOffset = findEmfRecord(original, 0x0e);
  const eof = emfRecord(0x0e, Buffer.alloc(20));
  eof.writeUInt32LE(2, 8);
  eof.writeUInt32LE(16, 12);
  eof.writeUInt32LE(0x00112233, 16);
  eof.writeUInt32LE(0x00445566, 20);
  eof.writeUInt32LE(eof.length, eof.length - 4);
  const result = Buffer.concat([original.subarray(0, eofOffset), eof]);
  result.writeUInt32LE(result.length, 48);
  return result;
}

function emfWithRestoredTextState(restoreValue: number): Buffer {
  const original = createRepresentativeEmf();
  const textOffset = findEmfRecord(original, 0x54);
  const fontOffset = findEmfRecord(original, 0x52);
  const changedFont = Buffer.from(
    original.subarray(fontOffset, fontOffset + original.readUInt32LE(fontOffset + 4)),
  );
  changedFont.writeUInt32LE(4, 8);
  changedFont.fill(0, 40, 104);
  Buffer.from("Changed Font", "utf16le").copy(changedFont, 40);
  const inserted = [
    emfRecord(0x21),
    emfRecord(0x18, uint32Values(0x0000ff00)),
    emfRecord(0x21),
    emfRecord(0x0a, int32Values(500, 500)),
    emfRecord(0x18, uint32Values(0x00ff0000)),
    emfRecord(0x16, uint32Values(6)),
    changedFont,
    emfRecord(0x25, uint32Values(4)),
    emfRecord(0x22, int32Values(restoreValue)),
  ];
  const result = Buffer.concat([
    original.subarray(0, textOffset),
    ...inserted,
    original.subarray(textOffset),
  ]);
  result.writeUInt32LE(result.length, 48);
  result.writeUInt32LE(result.readUInt32LE(52) + inserted.length, 52);
  return result;
}

function largeEmfRecord(type: number, recordSize: number): Buffer {
  const original = createRepresentativeEmf();
  const eofOffset = findEmfRecord(original, 0x0e);
  const record = Buffer.alloc(recordSize);
  record.writeUInt32LE(type, 0);
  record.writeUInt32LE(record.length, 4);
  const result = Buffer.concat([original.subarray(0, 88), record, original.subarray(eofOffset)]);
  result.writeUInt32LE(result.length, 48);
  result.writeUInt32LE(3, 52);
  return result;
}

function emfRecord(type: number, payload: Buffer = Buffer.alloc(0)): Buffer {
  const record = Buffer.alloc(8 + payload.length);
  record.writeUInt32LE(type, 0);
  record.writeUInt32LE(record.length, 4);
  payload.copy(record, 8);
  return record;
}

function int32Values(...values: number[]): Buffer {
  const result = Buffer.alloc(values.length * 4);
  values.forEach((value, index) => result.writeInt32LE(value, index * 4));
  return result;
}

function uint32Values(...values: number[]): Buffer {
  const result = Buffer.alloc(values.length * 4);
  values.forEach((value, index) => result.writeUInt32LE(value, index * 4));
  return result;
}
