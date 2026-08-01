import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  asOoxmlPercent,
  asPartPath,
  clearPictureCrop,
  countImageReferencesToMedia,
  readPptx,
  replaceImageBytes,
  setPictureCrop,
  type SourceImage,
  writePptx,
} from "../index.js";
import {
  BLUE_PNG,
  buildLayoutShowRoundTripFixture,
  buildMediaReplacementFixture,
  buildRoundTripFixture,
  buildUnreferencedLayoutFixture,
  decoder,
  encoder,
  getEntry,
  GREEN_PNG,
  RED_PNG,
} from "./write-pptx.test-helpers.js";

describe("writePptx - no-edit round-trip", () => {
  it("You can write no-edit PPTX from PptxSourceModel source and reload it.", () => {
    const input = buildRoundTripFixture();
    const original = readPptx(input);

    const output = writePptx(original);
    const reread = readPptx(output);

    expect(reread.presentation.slidePartPaths).toEqual(original.presentation.slidePartPaths);
    expect(reread.presentation.slideSize).toEqual(original.presentation.slideSize);
    expect(reread.slides.map((slide) => slide.partPath)).toEqual(
      original.slides.map((slide) => slide.partPath),
    );
  });

  it("Preserving relationship IDs and content types structurally", () => {
    const original = readPptx(buildRoundTripFixture());
    const reread = readPptx(writePptx(original));

    expect(reread.packageGraph.contentTypes).toEqual(original.packageGraph.contentTypes);
    expect(reread.packageGraph.relationships).toEqual(original.packageGraph.relationships);
  });

  it("Preserves p:sldLayout@show structurally in no-edit round-trip", () => {
    const input = buildLayoutShowRoundTripFixture();
    const source = readPptx(input);
    const output = writePptx(source);
    const reread = readPptx(output);

    expect(source.slideLayouts[0]?.show).toBe(false);
    expect(decoder.decode(getEntry(output, "ppt/slideLayouts/slideLayout1.xml"))).toContain(
      `show="0"`,
    );
    expect(reread.slideLayouts[0]?.show).toBe(false);
  });

  it("preserves master and layout id-list order in no-edit round-trip", () => {
    const original = readPptx(buildUnreferencedLayoutFixture());
    const reread = readPptx(writePptx(original));
    const originalMaster1 = original.slideMasters.find(
      (master) => master.partPath === "ppt/slideMasters/slideMaster1.xml",
    );
    const rereadMaster1 = reread.slideMasters.find(
      (master) => master.partPath === "ppt/slideMasters/slideMaster1.xml",
    );

    expect(original.presentation.slideMasterPartPaths).toEqual([
      "ppt/slideMasters/slideMaster2.xml",
      "ppt/slideMasters/slideMaster1.xml",
    ]);
    expect(reread.presentation.slideMasterPartPaths).toEqual(
      original.presentation.slideMasterPartPaths,
    );
    expect(originalMaster1?.layoutPartPaths).toEqual([
      "ppt/slideLayouts/slideLayout2.xml",
      "ppt/slideLayouts/slideLayout1.xml",
    ]);
    expect(rereadMaster1?.layoutPartPaths).toEqual(originalMaster1?.layoutPartPaths);
  });

  it("Preserving media bytes and unsupported raw package material", () => {
    const input = buildRoundTripFixture();
    const source = readPptx(input);
    const output = writePptx(source);

    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toBe(
      decoder.decode(getEntry(input, "docProps/custom.xml")),
    );
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).toBe(
      decoder.decode(getEntry(input, "ppt/slides/slide1.xml")),
    );
  });

  it("Replaces one image media part while preserving relationships, content types, and other media bytes", () => {
    const input = buildMediaReplacementFixture();
    const source = readPptx(input);
    const image = source.slides[0]?.shapes.find((shape) => shape.kind === "image");
    if (image === undefined) throw new Error("image fixture was not parsed");
    const edited = replaceImageBytes(source, image.handle!, BLUE_PNG);
    const output = writePptx(edited);
    const reread = readPptx(output);

    expect(getEntry(output, "ppt/media/image1.png")).toEqual(BLUE_PNG);
    expect(getEntry(output, "ppt/media/image2.png")).toEqual(GREEN_PNG);
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toBe(
      decoder.decode(getEntry(input, "docProps/custom.xml")),
    );
    expect(reread.packageGraph.contentTypes).toEqual(source.packageGraph.contentTypes);
    expect(reread.packageGraph.relationships).toEqual(source.packageGraph.relationships);
    expect(
      reread.packageGraph.media.find((part) => part.partPath === "ppt/media/image1.png"),
    ).toMatchObject({
      contentType: "image/png",
      bytes: BLUE_PNG,
    });
    expect(
      reread.packageGraph.media.find((part) => part.partPath === "ppt/media/image2.png"),
    ).toMatchObject({
      contentType: "image/png",
      bytes: GREEN_PNG,
    });
  });

  it("copy-on-write replaces only one of multiple pictures sharing a media part", () => {
    const input = buildMediaReplacementFixture(true);
    const source = readPptx(input);
    const target = source.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Replace Target",
    );
    if (target?.handle === undefined) throw new Error("shared image target was not parsed");

    const edited = replaceImageBytes(source, target.handle, BLUE_PNG);
    const output = writePptx(edited);
    const reread = readPptx(output);
    const targetAfter = reread.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Replace Target",
    );
    const sharedAfter = reread.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Keep Shared",
    );

    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "replaceImage",
      mode: "copyOnWrite",
      sourceMediaPartPath: "ppt/media/image1.png",
      mediaPartPath: "ppt/media/image3.png",
      replacementRelationshipId: "rId3",
      sharedReferenceCount: 2,
    });
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(getEntry(output, "ppt/media/image3.png")).toEqual(BLUE_PNG);
    expect(targetAfter?.blipRelationshipId).toBe("rId3");
    expect(sharedAfter?.blipRelationshipId).toBe("rIdImage1");
    expect(
      reread.packageGraph.relationships
        .find((group) => group.sourcePartPath === "ppt/slides/slide1.xml")
        ?.relationships.find((relationship) => relationship.id === "rId3"),
    ).toMatchObject({
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      target: "../media/image3.png",
    });
  });

  it("preserves the input relationship namespace prefix", () => {
    const source = readPptx(buildMediaReplacementFixture(true, "rel"));
    const target = source.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Replace Target",
    );
    if (target?.handle === undefined) throw new Error("prefixed image target was not parsed");

    const output = writePptx(replaceImageBytes(source, target.handle, BLUE_PNG));
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).toContain(`rel:embed="rId3"`);
    expect(slideXml).not.toContain(`r:embed="rId3"`);
  });

  it("updates the final reference in place after another shared picture is isolated", () => {
    const source = readPptx(buildMediaReplacementFixture(true));
    const target = source.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Replace Target",
    );
    const shared = source.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Keep Shared",
    );
    if (target?.handle === undefined || shared?.handle === undefined) {
      throw new Error("sequential image targets were not parsed");
    }

    const targetHandle = { partPath: target.handle.partPath, nodeId: target.handle.nodeId };
    const sharedHandle = { partPath: shared.handle.partPath, nodeId: shared.handle.nodeId };
    const isolated = replaceImageBytes(source, targetHandle, BLUE_PNG);
    expect(countImageReferencesToMedia(isolated, asPartPath("ppt/media/image1.png"))).toBe(1);
    const edited = replaceImageBytes(isolated, sharedHandle, GREEN_PNG);
    const output = writePptx(edited);

    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "replaceImage",
      mode: "inPlace",
      mediaPartPath: "ppt/media/image1.png",
      sharedReferenceCount: 1,
    });
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(GREEN_PNG);
    expect(getEntry(output, "ppt/media/image3.png")).toEqual(BLUE_PNG);
  });

  it("copy-on-write preserves a non-picture image fill that reuses the picture relationship", () => {
    const source = readPptx(buildMediaReplacementFixture("fill"));
    const target = source.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Replace Target",
    );
    if (target?.handle === undefined) throw new Error("image fill target was not parsed");

    const edited = replaceImageBytes(source, target.handle, BLUE_PNG);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "replaceImage",
      mode: "copyOnWrite",
      sharedReferenceCount: 2,
    });
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(getEntry(output, "ppt/media/image3.png")).toEqual(BLUE_PNG);
    expect(slideXml).toContain(`<a:blip r:embed="rId3"/>`);
    expect(slideXml).toContain(`<a:blip r:embed="rIdImage1"/>`);
  });

  it("copy-on-write preserves a VML image that uses r:id", () => {
    const source = readPptx(buildMediaReplacementFixture("vml"));
    const target = source.slides[0]?.shapes.find(
      (shape): shape is SourceImage => shape.kind === "image" && shape.name === "Replace Target",
    );
    if (target?.handle === undefined) throw new Error("VML image target was not parsed");

    const edited = replaceImageBytes(source, target.handle, BLUE_PNG);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "replaceImage",
      mode: "copyOnWrite",
      sharedReferenceCount: 2,
    });
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(slideXml).toContain(`<v:imagedata r:id="rIdImage1"/>`);
  });

  it("patches and clears only the targeted stretch picture srcRect, then rereads it", () => {
    const input = buildMediaReplacementFixture();
    const source = readPptx(input);
    const image = source.slides[0]?.shapes.find((shape) => shape.kind === "image");
    if (image?.handle === undefined) throw new Error("picture crop fixture was not parsed");
    expect(image.blipFillMode).toBe("stretch");

    const cropped = setPictureCrop(source, image.handle, {
      left: asOoxmlPercent(12000),
      top: asOoxmlPercent(3000),
      right: asOoxmlPercent(8000),
      bottom: asOoxmlPercent(4000),
    });
    const croppedOutput = writePptx(cropped);
    const croppedXml = decoder.decode(getEntry(croppedOutput, "ppt/slides/slide1.xml"));
    const reread = readPptx(croppedOutput);
    const rereadImage = reread.slides[0]?.shapes.find((shape) => shape.kind === "image");

    expect(croppedXml).toContain(`<a:srcRect x:l="preserve" l="12000" t="3000" r="8000" b="4000">`);
    expect(croppedXml).toContain(`data-preserve="yes"`);
    expect(croppedXml).toContain(`<x:keep value="yes"/>`);
    expect(croppedXml).toContain(`x:l="preserve"`);
    expect(croppedXml).toContain(`<x:inside value="yes"/>`);
    expect(croppedXml).toContain(`<a:stretch><a:fillRect/></a:stretch>`);
    expect(rereadImage?.crop).toEqual({ left: 12000, top: 3000, right: 8000, bottom: 4000 });

    const clearedOutput = writePptx(clearPictureCrop(cropped, image.handle));
    const clearedXml = decoder.decode(getEntry(clearedOutput, "ppt/slides/slide1.xml"));
    const clearedImage = readPptx(clearedOutput).slides[0]?.shapes.find(
      (shape) => shape.kind === "image",
    );
    expect(clearedXml).not.toContain("srcRect");
    expect(clearedXml).toContain(`<x:keep value="yes"/>`);
    expect(clearedImage?.crop).toBeUndefined();
  });

  it("uses the existing DrawingML prefix when inserting picture srcRect", () => {
    const entries = unzipSync(buildMediaReplacementFixture());
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = decoder
      .decode(entries[slidePath])
      .replace(`<a:srcRect l="1000" x:l="preserve"><x:inside value="yes"/></a:srcRect>`, "")
      .replace(
        `<a:stretch><a:fillRect/></a:stretch>`,
        `<d:stretch xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fillRect/></d:stretch>`,
      );
    entries[slidePath] = encoder.encode(slideXml);
    const source = readPptx(zipSync(entries));
    const image = source.slides[0]?.shapes.find((shape) => shape.kind === "image");
    if (image?.handle === undefined) throw new Error("prefixed picture was not parsed");

    const output = writePptx(setPictureCrop(source, image.handle, { left: asOoxmlPercent(10000) }));
    const writtenXml = decoder.decode(getEntry(output, slidePath));

    expect(writtenXml).toContain(
      `<d:srcRect l="10000" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
    );
    expect(writtenXml).not.toContain(`<a:srcRect`);
  });

  it("rejects duplicate picture crop edits even when relationship ids differ", () => {
    const source = readPptx(buildMediaReplacementFixture());
    const image = source.slides[0]?.shapes.find((shape) => shape.kind === "image");
    if (image?.handle === undefined) throw new Error("picture crop fixture was not parsed");
    const conflicted = {
      ...source,
      edits: [
        {
          kind: "updatePictureCrop" as const,
          handle: image.handle,
          crop: { left: asOoxmlPercent(10000) },
        },
        {
          kind: "updatePictureCrop" as const,
          handle: { ...image.handle, relationshipId: undefined },
          crop: { right: asOoxmlPercent(10000) },
        },
      ],
    };

    expect(() => writePptx(conflicted)).toThrow(/conflicting picture crop edits/);
  });

  it("Can write xml raw package material in serializable range", () => {
    const source = readPptx(buildRoundTripFixture());
    const withXmlRaw = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        contentTypes: {
          ...source.packageGraph.contentTypes,
          overrides: [
            ...source.packageGraph.contentTypes.overrides,
            { partName: "customXml/item1.xml", contentType: "application/xml" },
          ],
        },
        parts: [
          ...source.packageGraph.parts,
          { partPath: "customXml/item1.xml", contentType: "application/xml" },
        ],
        rawParts: [
          ...(source.packageGraph.rawParts ?? []),
          {
            kind: "xml",
            partPath: "customXml/item1.xml",
            contentType: "application/xml",
            xml: {
              name: "x:root",
              attributes: { "xmlns:x": "urn:test", "a:flag": "A&B" },
              children: [{ name: "x:child", text: "nested < text" }],
            },
          },
        ],
      },
    };

    expect(decoder.decode(getEntry(writePptx(withXmlRaw), "customXml/item1.xml"))).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<x:root xmlns:x="urn:test" a:flag="A&amp;B"><x:child>nested &lt; text</x:child></x:root>`,
    );
  });

  it("Reject mixed content of xml raw package material because order cannot be maintained", () => {
    const source = readPptx(buildRoundTripFixture());
    const withMixedContentRaw = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        parts: [
          ...source.packageGraph.parts,
          { partPath: "customXml/item1.xml", contentType: "application/xml" },
        ],
        rawParts: [
          ...(source.packageGraph.rawParts ?? []),
          {
            kind: "xml",
            partPath: "customXml/item1.xml",
            contentType: "application/xml",
            xml: {
              name: "x:root",
              text: "pre",
              children: [{ name: "x:child" }],
            },
          },
        ],
      },
    };

    expect(() => writePptx(withMixedContentRaw)).toThrow(/mixed text\/element content/);
  });

  it("Verify structural preservation rather than byte equality", () => {
    const input = buildRoundTripFixture();
    const output = writePptx(readPptx(input));

    expect(output).not.toEqual(input);
    expect(readPptx(output).presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
  });

  it("Even in real fixtures, PPTX after write can be reread with readPptx", () => {
    const fixturePath = fileURLToPath(
      new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
    );
    const source = readPptx(readFileSync(fixturePath));
    const output = writePptx(source);
    const reread = readPptx(output);
    const originalImage = source.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );
    const rereadImage = reread.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );

    expect(reread.presentation.slidePartPaths).toEqual(source.presentation.slidePartPaths);
    expect(reread.presentation.slideSize).toEqual(source.presentation.slideSize);
    expect(rereadImage?.bytes).toEqual(originalImage?.bytes);
  });

  it("Parts without preserved material are not implicitly regenerated by no-edit writer.", () => {
    const source = readPptx(buildRoundTripFixture());
    const withoutRawSlide = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        rawParts: source.packageGraph.rawParts?.filter(
          (part) => part.partPath !== "ppt/slides/slide1.xml",
        ),
      },
    };

    expect(() => writePptx(withoutRawSlide)).toThrow(/no preserved package material/);
  });
});
