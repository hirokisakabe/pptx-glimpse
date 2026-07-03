import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import { asEmu, asSourceNodeId, readPptx, writePptx } from "../index.js";
import {
  findParagraphBySourceHandle,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
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
