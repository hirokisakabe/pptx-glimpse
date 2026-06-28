import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Test note.
import { readPptx, writePptx } from "../index.js";
import { findTextRunBySourceHandle, replaceTextRunPlainText, type SourceShape } from "../index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function buildRoundTripFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst>` +
        `<p:sldId id="256" r:id="rIdSlide1"/>` +
        `<p:sldId id="257" r:id="rIdSlide2"/>` +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/slide2.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>` +
        `</Relationships>`,
    ),
    "ppt/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
    "docProps/custom.xml": xml(`<Properties><custom value="preserve-me"/></Properties>`),
  });
}

function buildTextEditFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>` +
      `<a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1" i="1" sz="2400"><a:latin typeface="Aptos"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ea typeface="Noto Sans JP"/></a:rPr><a:t>Original</a:t></a:r><a:r><a:t xml:space="preserve"> Keep </a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`,
  );
}

function buildSlotHandleTextEditFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>` +
      `<p:sp><p:nvSpPr><p:cNvPr name="No Id Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original</a:t></a:r></a:p></p:txBody></p:sp>`,
  );
}

function buildTextEditFixtureFromSlide(slideSpTree: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>${slideSpTree}</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
    "docProps/custom.xml": xml(`<Properties><custom value="preserve-me"/></Properties>`),
    "ppt/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9, 8, 7]),
  });
}

function getEntry(output: Uint8Array, path: string): Uint8Array {
  const entry = unzipSync(output)[path];
  if (entry === undefined) throw new Error(`missing zip entry: ${path}`);
  return entry;
}

describe("writePptx — no-edit round-trip", () => {
  it("covers write-pptx behavior 1", () => {
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

  it("covers write-pptx behavior 2", () => {
    const original = readPptx(buildRoundTripFixture());
    const reread = readPptx(writePptx(original));

    expect(reread.packageGraph.contentTypes).toEqual(original.packageGraph.contentTypes);
    expect(reread.packageGraph.relationships).toEqual(original.packageGraph.relationships);
  });

  it("covers write-pptx behavior 3", () => {
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

  it("covers write-pptx behavior 4", () => {
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

  it("covers write-pptx behavior 5", () => {
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

  it("covers write-pptx behavior 6", () => {
    const input = buildRoundTripFixture();
    const output = writePptx(readPptx(input));

    expect(output).not.toEqual(input);
    expect(readPptx(output).presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
  });

  it("covers write-pptx behavior 7", () => {
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

  it("covers write-pptx behavior 8", () => {
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

describe("writePptx — one plain text-run edit", () => {
  it("covers write-pptx behavior 9", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shape:10:p:0:r:0",
      orderingSlot: 0,
    });
    expect(findTextRunBySourceHandle(source, run.handle!)).toBe(run);
  });

  it("covers write-pptx behavior 10", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const run = firstRun(source);

    const edited = replaceTextRunPlainText(source, run.handle!, "Edited text");
    const reread = readPptx(writePptx(edited));
    const editedRun = firstRun(reread);

    expect(run.text).toBe("Original");
    expect(firstRun(edited).text).toBe("Edited text");
    expect(editedRun.text).toBe("Edited text");
    expect(firstParagraph(reread).runs[1].text).toBe(" Keep ");
  });

  it("covers write-pptx behavior 11", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const edited = replaceTextRunPlainText(source, firstRun(source).handle!, "Edited text");
    const output = writePptx(edited);
    const reread = readPptx(output);
    const run = firstRun(reread);
    const paragraph = firstParagraph(reread);

    expect(paragraph.properties).toEqual({ align: "center" });
    expect(run.properties).toMatchObject({
      bold: true,
      italic: true,
      fontSize: 24,
      typeface: "Aptos",
      typefaceEa: "Noto Sans JP",
      color: { kind: "srgb", hex: "FF0000" },
    });
    expect(run.rawSidecars?.map((sidecar) => sidecar.node.name) ?? []).not.toContain("a:ea");
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("covers write-pptx behavior 12", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shapeSlot:1:p:0:r:0",
      orderingSlot: 0,
    });

    const edited = replaceTextRunPlainText(source, run.handle!, "Edited via slot");
    expect(firstRun(readPptx(writePptx(edited))).text).toBe("Edited via slot");
  });
});

function firstShape(source: ReturnType<typeof readPptx>): SourceShape {
  const shape = source.slides[0].shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("shape not found");
  return shape;
}

function firstParagraph(source: ReturnType<typeof readPptx>) {
  return firstShape(source).textBody!.paragraphs[0];
}

function firstRun(source: ReturnType<typeof readPptx>) {
  return firstParagraph(source).runs[0];
}
