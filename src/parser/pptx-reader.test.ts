import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { LazyMediaMap, LazyXmlMap, readPptx } from "./pptx-reader.js";

function createTestZip(entries: Record<string, Uint8Array>): Uint8Array {
  return zipSync(entries);
}

describe("LazyMediaMap", () => {
  it("returns media data for an existing path", () => {
    const imageData = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    const zip = createTestZip({
      "ppt/media/image1.png": imageData,
      "ppt/presentation.xml": strToU8("<xml/>"),
    });
    const entryIndex = new Set(["ppt/media/image1.png"]);
    const media = new LazyMediaMap(zip, entryIndex);

    const result = media.get("ppt/media/image1.png");
    expect(result).toEqual(imageData);
  });

  it("returns undefined for a non-existent path without calling unzipSync", () => {
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8("<xml/>"),
    });
    const entryIndex = new Set<string>();
    const media = new LazyMediaMap(zip, entryIndex);

    const result = media.get("ppt/media/nonexistent.png");
    expect(result).toBeUndefined();
  });

  it("caches the result for repeated access", () => {
    const imageData = new Uint8Array([0xff, 0xd8, 0xff, 0xe0]);
    const zip = createTestZip({
      "ppt/media/photo.jpg": imageData,
    });
    const entryIndex = new Set(["ppt/media/photo.jpg"]);
    const media = new LazyMediaMap(zip, entryIndex);

    const first = media.get("ppt/media/photo.jpg");
    const second = media.get("ppt/media/photo.jpg");
    expect(first).toBe(second);
  });
});

describe("LazyXmlMap", () => {
  it("returns XML string for an existing path", () => {
    const xmlContent = "<p:presentation/>";
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8(xmlContent),
    });
    const entryIndex = new Set(["ppt/presentation.xml"]);
    const xmlMap = new LazyXmlMap(zip, entryIndex);

    expect(xmlMap.get("ppt/presentation.xml")).toBe(xmlContent);
  });

  it("returns undefined for a non-existent path", () => {
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8("<xml/>"),
    });
    const entryIndex = new Set<string>();
    const xmlMap = new LazyXmlMap(zip, entryIndex);

    expect(xmlMap.get("ppt/slides/slide99.xml")).toBeUndefined();
  });

  it("has() returns true for indexed paths and false otherwise", () => {
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8("<xml/>"),
    });
    const entryIndex = new Set(["ppt/presentation.xml"]);
    const xmlMap = new LazyXmlMap(zip, entryIndex);

    expect(xmlMap.has("ppt/presentation.xml")).toBe(true);
    expect(xmlMap.has("ppt/slides/slide1.xml")).toBe(false);
  });

  it("caches the result for repeated access", () => {
    const xmlContent = "<p:presentation/>";
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8(xmlContent),
    });
    const entryIndex = new Set(["ppt/presentation.xml"]);
    const xmlMap = new LazyXmlMap(zip, entryIndex);

    const first = xmlMap.get("ppt/presentation.xml");
    const second = xmlMap.get("ppt/presentation.xml");
    // Cached string is the same object reference (not re-decoded)
    expect(first).toBe(second);
  });
});

describe("readPptx", () => {
  it("provides lazy access to XML and media files", () => {
    const xmlContent = "<p:presentation/>";
    const imageData = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8(xmlContent),
      "ppt/_rels/presentation.xml.rels": strToU8("<Relationships/>"),
      "ppt/media/image1.png": imageData,
    });

    const archive = readPptx(zip);

    expect(archive.files.get("ppt/presentation.xml")).toBe(xmlContent);
    expect(archive.files.get("ppt/_rels/presentation.xml.rels")).toBe("<Relationships/>");
    expect(archive.media.get("ppt/media/image1.png")).toEqual(imageData);
  });

  it("does not include media files in the XML accessor", () => {
    const imageData = new Uint8Array(1024).fill(0xab);
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8("<xml/>"),
      "ppt/media/large-image.png": imageData,
    });

    const archive = readPptx(zip);

    expect(archive.files.has("ppt/media/large-image.png")).toBe(false);
    expect(archive.media.get("ppt/media/large-image.png")).toEqual(imageData);
  });

  it("returns undefined for XML files not in the archive", () => {
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8("<xml/>"),
    });

    const archive = readPptx(zip);
    expect(archive.files.get("ppt/slides/slide99.xml")).toBeUndefined();
  });

  it("returns undefined for non-existent media", () => {
    const zip = createTestZip({
      "ppt/presentation.xml": strToU8("<xml/>"),
    });

    const archive = readPptx(zip);
    expect(archive.media.get("ppt/media/nonexistent.png")).toBeUndefined();
  });
});
