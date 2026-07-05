import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import { asEmu, asPt, asSourceNodeId, readPptx, writePptx } from "../index.js";
import {
  clearTextRunProperties,
  deleteSlide,
  duplicateSlide,
  findParagraphBySourceHandle,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setTextRunProperties,
  type SourceShape,
  updateShapeTransform,
} from "../index.js";

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

function buildSlideTopologyFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` +
        `<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>` +
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
        `<p:sldId id="300" r:id="rIdSlide2"/>` +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Invisible Source"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slide One</a:t></a:r></a:p></p:txBody>` +
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `<p:timing><p:tnLst><p:par/></p:tnLst></p:timing>` +
        `</p:sld>`,
    ),
    "ppt/slides/slide2.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="20" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slide Two</a:t></a:r></a:p></p:txBody>` +
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>` +
        `<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/notesSlides/notesSlide1.xml": xml(
      `<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:notes>`,
    ),
    "ppt/notesSlides/_rels/notesSlide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/>` +
        `<Relationship Id="rIdNotesMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/comments/comment1.xml": xml(
      `<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
    ),
  });
}

function buildSlideTopologyFixtureWithRelationshipOverrides(): Uint8Array {
  const entries = unzipSync(buildSlideTopologyFixture());
  const relationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";
  const contentTypes = decoder
    .decode(entries["[Content_Types].xml"])
    .replace(`<Default Extension="rels" ContentType="${relationshipContentType}"/>`, "")
    .replace(
      `</Types>`,
      `<Override PartName="/_rels/.rels" ContentType="${relationshipContentType}"/>` +
        `<Override PartName="/ppt/_rels/presentation.xml.rels" ContentType="${relationshipContentType}"/>` +
        `<Override PartName="/ppt/slides/_rels/slide1.xml.rels" ContentType="${relationshipContentType}"/>` +
        `<Override PartName="/ppt/notesSlides/_rels/notesSlide1.xml.rels" ContentType="${relationshipContentType}"/>` +
        `</Types>`,
    );

  return zipSync({
    ...entries,
    "[Content_Types].xml": encoder.encode(contentTypes),
  });
}

function buildLayoutShowRoundTripFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
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
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": xml(
      `<p:sldLayout show="0" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster1.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldMaster>`,
    ),
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

function buildMultipleTextEditFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>First paragraph</a:t></a:r></a:p>` +
      `<a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="11" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>Other shape</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`,
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

describe("writePptx - slide topology edits", () => {
  it("duplicates a slide immediately after the source and preserves raw invisible slide material", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = duplicateSlide(source, source.slides[0].handle!);
    const output = writePptx(edited);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));
    const duplicateSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide3.xml"));
    const duplicateSlideRels = decoder.decode(getEntry(output, "ppt/slides/_rels/slide3.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(
      reread.slides.map((slide) => slide.shapes[0]?.kind === "shape" && slide.shapes[0].name),
    ).toEqual(["Invisible Source", "Invisible Source", "Second"]);
    expect(presentationXml).toContain(`<p:sldId id="301" r:id="rId10"/>`);
    expect(presentationXml.indexOf(`r:id="rIdSlide1"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rId10"`),
    );
    expect(presentationXml.indexOf(`r:id="rId10"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rIdSlide2"`),
    );
    expect(presentationRels).toContain(
      `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>`,
    );
    expect(duplicateSlideXml).toContain("<p:timing>");
    expect(duplicateSlideRels).toContain(`Id="rIdComments"`);
    expect(duplicateSlideRels).toContain(`Target="../comments/comment1.xml"`);
    expect(duplicateSlideRels).toContain(`Id="rIdNotes"`);
    expect(duplicateSlideRels).toContain(`Target="../notesSlides/notesSlide2.xml"`);
    expect(decoder.decode(getEntry(output, "ppt/notesSlides/notesSlide2.xml"))).toContain(
      "<p:notes",
    );
    expect(
      decoder.decode(getEntry(output, "ppt/notesSlides/_rels/notesSlide2.xml.rels")),
    ).toContain(`notesMaster`);
    expect(
      decoder.decode(getEntry(output, "ppt/notesSlides/_rels/notesSlide2.xml.rels")),
    ).toContain(`Target="../slides/slide3.xml"`);
    expect(reread.packageGraph.contentTypes.overrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/slide3.xml",
          contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        },
        {
          partName: "ppt/notesSlides/notesSlide2.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        },
      ]),
    );
  });

  it("deletes a slide and its notes part while keeping remaining slide order and orphan cleanup out of scope", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = deleteSlide(source, source.slides[0].handle!);
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
    expect(reread.slides[0]?.shapes[0]?.kind === "shape" && reread.slides[0].shapes[0].name).toBe(
      "Second",
    );
    expect(presentationXml).not.toContain(`r:id="rIdSlide1"`);
    expect(presentationRels).not.toContain(`Id="rIdSlide1"`);
    expect(entries["ppt/slides/slide1.xml"]).toBeUndefined();
    expect(entries["ppt/slides/_rels/slide1.xml.rels"]).toBeUndefined();
    expect(entries["ppt/notesSlides/notesSlide1.xml"]).toBeUndefined();
    expect(entries["ppt/notesSlides/_rels/notesSlide1.xml.rels"]).toBeUndefined();
    expect(entries["ppt/comments/comment1.xml"]).toBeDefined();
    expect(
      reread.packageGraph.contentTypes.overrides.some(
        (override) => override.partName === "ppt/notesSlides/notesSlide1.xml",
      ),
    ).toBe(false);
  });

  it("keeps relationship content type overrides consistent when no rels default exists", () => {
    const source = readPptx(buildSlideTopologyFixtureWithRelationshipOverrides());
    const duplicated = readPptx(writePptx(duplicateSlide(source, source.slides[0].handle!)));
    const deleted = readPptx(writePptx(deleteSlide(source, source.slides[0].handle!)));
    const duplicatedOverrides = duplicated.packageGraph.contentTypes.overrides;
    const deletedOverridePartNames = new Set(
      deleted.packageGraph.contentTypes.overrides.map((override) => override.partName),
    );

    expect(
      duplicated.packageGraph.contentTypes.defaults.some((entry) => entry.extension === "rels"),
    ).toBe(false);
    expect(duplicatedOverrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/_rels/slide3.xml.rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
        {
          partName: "ppt/notesSlides/_rels/notesSlide2.xml.rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
      ]),
    );
    expect(deletedOverridePartNames.has("ppt/slides/slide1.xml")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/slides/_rels/slide1.xml.rels")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/notesSlides/notesSlide1.xml")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/notesSlides/_rels/notesSlide1.xml.rels")).toBe(false);
  });

  it("rejects deleting the last slide and duplicating a dirty slide", () => {
    const singleSlide = readPptx(buildTextEditFixture());
    expect(() => deleteSlide(singleSlide, singleSlide.slides[0].handle!)).toThrow(/last slide/);

    const dirty = replaceTextRunPlainText(singleSlide, firstRun(singleSlide).handle!, "Dirty");
    expect(() => duplicateSlide(dirty, dirty.slides[0].handle!)).toThrow(
      /pending dirty part edits/,
    );
  });

  it("does not reuse a deleted slide part name within the same edit journal", () => {
    const source = readPptx(buildRoundTripFixture());
    const deletedSecond = deleteSlide(source, source.slides[1].handle!);
    const duplicatedFirst = duplicateSlide(deletedSecond, deletedSecond.slides[0].handle!);
    const deletedDuplicate = deleteSlide(duplicatedFirst, duplicatedFirst.slides[1].handle!);
    const output = writePptx(deletedDuplicate);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));

    expect(duplicatedFirst.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
    ]);
    expect(reread.presentation.slidePartPaths).toEqual(["ppt/slides/slide1.xml"]);
    expect(presentationXml).not.toContain(`r:id="rIdSlide2"`);
    expect(presentationXml).not.toContain(`r:id="rId3"`);
  });
});

describe("writePptx - one plain text-run edit", () => {
  it("Existing text run can be identified with stable source handle", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shape:10:p:0:r:0",
      orderingSlot: 0,
    });
    expect(findTextRunBySourceHandle(source, run.handle!)).toBe(run);
  });

  it("Apply plain text replacement to PptxSourceModel source and reflect in PPTX after write", () => {
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

  it("Writes dirty slide XML with a single XML declaration", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = replaceTextRunPlainText(source, firstRun(source).handle!, "Edited text");
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(slideXml.match(/<\?xml/g)).toHaveLength(1);
  });

  it("Preserving run / paragraph formatting and unrelated package material", () => {
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

  it("Runs without shape ids can also be reflected in dirty slide XML using shapeSlot handle.", () => {
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

describe("writePptx - multiple text edits", () => {
  it("Applies multiple text-run replacements across different shapes and paragraphs in one write.", () => {
    const source = readPptx(buildMultipleTextEditFixture());
    const firstShapeSecondParagraphRun = shapeAt(source, 0).textBody!.paragraphs[1].runs[0];
    const secondShapeRun = shapeAt(source, 1).textBody!.paragraphs[0].runs[0];

    const edited = replaceTextRunPlainText(
      replaceTextRunPlainText(source, firstShapeSecondParagraphRun.handle!, "Edited paragraph"),
      secondShapeRun.handle!,
      "Edited other shape",
    );
    const reread = readPptx(writePptx(edited));

    expect(shapeAt(reread, 0).textBody!.paragraphs[0].runs[0].text).toBe("First paragraph");
    expect(shapeAt(reread, 0).textBody!.paragraphs[1].runs[0].text).toBe("Edited paragraph");
    expect(shapeAt(reread, 1).textBody!.paragraphs[0].runs[0].text).toBe("Edited other shape");
  });

  it("Rejects conflicting duplicate edits for the same text run.", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);
    const edited = replaceTextRunPlainText(
      replaceTextRunPlainText(source, run.handle!, "First edit"),
      run.handle!,
      "Second edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run edits/);
  });

  it("Rejects conflicting text-run and paragraph replacements for the same paragraph.", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);
    const edited = replaceParagraphPlainText(
      replaceTextRunPlainText(source, paragraph.runs[0].handle!, "Run edit"),
      paragraph.handle!,
      "Paragraph edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run and paragraph edits/);
  });
});

describe("writePptx - text run property edits", () => {
  it("Sets all supported text run properties and persists them after write/read", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    const edited = setTextRunProperties(source, run.handle!, {
      bold: false,
      italic: false,
      underline: true,
      fontSize: asPt(32),
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Liberation Sans",
    });
    const reread = readPptx(writePptx(edited));
    const editedRun = firstRun(reread);

    expect(firstRun(edited).properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 32,
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Liberation Sans",
    });
    expect(editedRun.properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 32,
      color: { kind: "srgb", hex: "00AA44" },
      typeface: "Liberation Sans",
      typefaceEa: "Noto Sans JP",
    });
  });

  it("Clears supported text run properties without removing unrelated rPr attributes or children", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="70" name="Decorated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr lang="en-US" b="1" i="1" u="sng" sz="2400" spc="120">` +
          `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill>` +
          `<a:latin typeface="Aptos"/><a:effectLst/><a:ea typeface="Noto Sans JP"/>` +
          `</a:rPr><a:t>Original</a:t></a:r>` +
          `<a:r><a:rPr b="1"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>Untouched</a:t></a:r>` +
          `</a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    const edited = clearTextRunProperties(source, firstRun(source).handle!, [
      "bold",
      "italic",
      "underline",
      "fontSize",
      "color",
      "typeface",
    ]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    const properties = firstRun(reread).properties;
    expect(properties).toMatchObject({ typefaceEa: "Noto Sans JP" });
    expect(properties?.bold).toBeUndefined();
    expect(properties?.italic).toBeUndefined();
    expect(properties?.underline).toBeUndefined();
    expect(properties?.fontSize).toBeUndefined();
    expect(properties?.color).toBeUndefined();
    expect(properties?.typeface).toBeUndefined();
    expect(slideXml).toContain('lang="en-US"');
    expect(slideXml).toContain('spc="120"');
    expect(slideXml).toContain("<a:effectLst");
    expect(slideXml).toContain('<a:ea typeface="Noto Sans JP"');
    expect(firstParagraph(reread).runs[1].properties).toMatchObject({
      bold: true,
      color: { kind: "srgb", hex: "00FF00" },
    });
  });

  it("Creates rPr when setting properties on a run without existing properties", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const edited = setTextRunProperties(source, firstRun(source).handle!, {
      bold: true,
      fontSize: asPt(18),
      color: { kind: "srgb", hex: "112233" },
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml.indexOf("<a:rPr")).toBeLessThan(slideXml.indexOf("<a:t>Original</a:t>"));
    expect(firstRun(readPptx(output)).properties).toMatchObject({
      bold: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "112233" },
    });
  });

  it("Does not dirty a run when clearing properties that are already absent", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const edited = clearTextRunProperties(source, firstRun(source).handle!, ["bold"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(edited).toBe(source);
    expect(slideXml).not.toContain("<a:rPr");
    expect(firstRun(readPptx(output)).properties).toBeUndefined();
  });

  it("Removes an empty latin run property element when clearing typeface", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="71" name="Typeface"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr><a:latin typeface="Aptos"/></a:rPr><a:t>Original</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const edited = clearTextRunProperties(source, firstRun(source).handle!, ["typeface"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).not.toContain("<a:latin");
    expect(slideXml).not.toContain("<a:rPr");
    expect(firstRun(readPptx(output)).properties).toBeUndefined();
  });

  it("Rejects invalid direct text run property helper input", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstRun(source).handle!;

    expect(() => setTextRunProperties(source, handle, { fontSize: asPt(0) })).toThrow(
      /fontSize must be a finite positive pt value/,
    );
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      setTextRunProperties(source, handle, { strikethrough: true }),
    ).toThrow(/unsupported text run property 'strikethrough'/);
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      clearTextRunProperties(source, handle, ["strikethrough"]),
    ).toThrow(/unsupported text run property 'strikethrough'/);
  });

  it("Rejects no-op text run property edits constructed directly in an edit journal", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = {
      ...source,
      edits: [
        {
          kind: "updateTextRunProperties",
          handle: firstRun(source).handle!,
        },
      ],
    } satisfies typeof source;

    expect(() => writePptx(edited)).toThrow(/must set or clear at least one property/);
  });

  it("Applies text and property edits to the same run in one write", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstRun(source).handle!;
    const edited = setTextRunProperties(
      replaceTextRunPlainText(source, handle, "Edited property text"),
      handle,
      { underline: true, color: { kind: "srgb", hex: "336699" } },
    );
    const reread = readPptx(writePptx(edited));

    expect(firstRun(reread).text).toBe("Edited property text");
    expect(firstRun(reread).properties).toMatchObject({
      underline: true,
      color: { kind: "srgb", hex: "336699" },
    });
  });

  it("Rejects conflicting text run property and paragraph replacements for the same paragraph", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);
    const edited = replaceParagraphPlainText(
      setTextRunProperties(source, paragraph.runs[0].handle!, { bold: false }),
      paragraph.handle!,
      "Paragraph edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run properties and paragraph edits/);
  });
});

describe("writePptx - paragraph text replacement", () => {
  it("Normalizes a multi-run paragraph to one run using the first run properties.", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const paragraph = firstParagraph(source);

    expect(findParagraphBySourceHandle(source, paragraph.handle!)).toBe(paragraph);

    const edited = replaceParagraphPlainText(source, paragraph.handle!, "Paragraph replacement");
    const output = writePptx(edited);
    const reread = readPptx(output);
    const editedParagraph = firstParagraph(reread);

    expect(firstParagraph(edited).runs).toHaveLength(1);
    expect(editedParagraph.runs).toHaveLength(1);
    expect(editedParagraph.runs[0].text).toBe("Paragraph replacement");
    expect(editedParagraph.runs[0].properties).toMatchObject({
      bold: true,
      italic: true,
      fontSize: 24,
      typeface: "Aptos",
      typefaceEa: "Noto Sans JP",
      color: { kind: "srgb", hex: "FF0000" },
    });
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).not.toContain(" Keep ");
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Round-trips paragraph replacement text with significant surrounding whitespace.", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = replaceParagraphPlainText(source, firstParagraph(source).handle!, " Trimmed ");
    const reread = readPptx(writePptx(edited));

    expect(firstParagraph(reread).runs).toHaveLength(1);
    expect(firstRun(reread).text).toBe(" Trimmed ");
  });

  it("Keeps replacement run before endParaRPr.", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="60" name="End para props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr algn="ctr"/><a:r><a:t>Before</a:t></a:r><a:endParaRPr lang="ja-JP"/></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const output = writePptx(
      replaceParagraphPlainText(source, firstParagraph(source).handle!, "After"),
    );
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml.indexOf("<a:r>")).toBeGreaterThan(-1);
    expect(slideXml.indexOf("<a:endParaRPr")).toBeGreaterThan(-1);
    expect(slideXml.indexOf("<a:r>")).toBeLessThan(slideXml.indexOf("<a:endParaRPr"));
    expect(firstRun(readPptx(output)).text).toBe("After");
  });

  it("Rejects paragraph replacement for interleaved bullet paragraphs split by the reader.", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="61" name="Interleaved bullets"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p>` +
          `<a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>One</a:t></a:r><a:br/>` +
          `<a:pPr><a:buChar char="&#x25E6;"/></a:pPr><a:r><a:t>Two</a:t></a:r>` +
          `</a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    expect(firstShape(source).textBody!.paragraphs).toHaveLength(2);
    const edited = replaceParagraphPlainText(source, firstParagraph(source).handle!, "After");

    expect(() => writePptx(edited)).toThrow(/interleaved bullet paragraph/);
  });
});

describe("writePptx - shape xfrm edit", () => {
  it("Apply offset and extent update to PptxSourceModel source and reflect in PPTX after write", () => {
    const source = readPptx(buildTextEditFixture());
    const shape = firstShape(source);
    const handle = shape.handle!;

    const edited = updateShapeTransform(source, handle, {
      offsetX: asEmu(1111),
      offsetY: asEmu(2222),
      width: asEmu(3333),
      height: asEmu(4444),
    });
    const reread = readPptx(writePptx(edited));
    const editedShape = findShapeNodeBySourceHandle(reread, handle);

    expect(firstShape(edited).transform).toMatchObject({
      offsetX: 1111,
      offsetY: 2222,
      width: 3333,
      height: 4444,
    });
    expect(editedShape?.transform).toMatchObject({
      offsetX: 1111,
      offsetY: 2222,
      width: 3333,
      height: 4444,
    });
  });

  it("Preserves unrelated package material while replacing only dirty slide XML", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const edited = updateShapeTransform(source, firstShape(source).handle!, {
      offsetX: asEmu(1111),
      offsetY: asEmu(2222),
      width: asEmu(3333),
      height: asEmu(4444),
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).toContain('<a:off x="1111" y="2222"');
    expect(slideXml).toContain('<a:ext cx="3333" cy="4444"');
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Rejects shape handles that do not have xfrm", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="50" name="No xfrm"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`,
      ),
    );
    const shapeWithoutXfrm = firstShape(source);

    expect(() =>
      updateShapeTransform(source, shapeWithoutXfrm.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/does not reference a shape with xfrm/);
  });

  it("Rejects shape transform handles without node ids", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr name="No Id Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`,
      ),
    );
    const shape = firstShape(source);

    expect(shape.handle?.nodeId).toBeUndefined();
    expect(() =>
      updateShapeTransform(source, shape.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/requires a node id/);
  });

  it("Rejects nonexistent shape handles", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = {
      ...firstShape(source).handle!,
      nodeId: asSourceNodeId("999"),
    };

    expect(() =>
      updateShapeTransform(source, handle, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/shape handle was not found/);
  });

  it("Rejects nested group child shape handles for this writer slice", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="30" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/><a:chOff x="0" y="0"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="31" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `</p:sp>` +
          `</p:grpSp>`,
      ),
    );
    const group = source.slides[0].shapes[0];
    if (group.kind !== "group") throw new Error("group not found");
    const child = group.children[0];

    expect(findShapeNodeBySourceHandle(source, child.handle!)).toBe(child);
    expect(() =>
      updateShapeTransform(source, child.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/nested group shape editing is not supported/);
  });

  it("Rejects AlternateContent fallback shape handles for this writer slice", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Fallback>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="40" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `</p:sp>` +
          `</mc:Fallback>` +
          `</mc:AlternateContent>`,
      ),
    );
    const shape = firstShape(source);

    expect(() =>
      updateShapeTransform(source, shape.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/AlternateContent/);
  });
});

function firstShape(source: ReturnType<typeof readPptx>): SourceShape {
  return shapeAt(source, 0);
}

function shapeAt(source: ReturnType<typeof readPptx>, index: number): SourceShape {
  const shape = source.slides[0].shapes.filter(
    (node): node is SourceShape => node.kind === "shape",
  )[index];
  if (shape === undefined) throw new Error("shape not found");
  return shape;
}

function firstParagraph(source: ReturnType<typeof readPptx>) {
  return firstShape(source).textBody!.paragraphs[0];
}

function firstRun(source: ReturnType<typeof readPptx>) {
  return firstParagraph(source).runs[0];
}
