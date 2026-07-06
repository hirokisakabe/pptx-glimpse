import { Buffer } from "node:buffer";

import { asEmu, readPptx, type SourceConnector, type SourceShape } from "@pptx-glimpse/document";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { createBrowserPptxEditorSession } from "./browser-editor.js";

const encoder = new TextEncoder();
const RED_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=",
);
const BLUE_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
);

describe("BrowserPptxEditorSession", () => {
  it("edits, renders, undoes, redoes, and saves a browser editor session", async () => {
    const editor = await createBrowserPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const shape = editor.shapes(1)[0];
    if (shape?.handle === undefined) throw new Error("shape handle not found");

    expect(editor.slides).toHaveLength(1);
    expect(shape.bounds).toEqual({ x: 96, y: 192, width: 288, height: 96 });
    expect(shape.editableTextBody).toBeDefined();

    await editor.apply({
      kind: "setShapeTransform",
      handle: shape.handle,
      offsetX: asEmu(120 * 9525),
      offsetY: asEmu(208 * 9525),
      width: asEmu(336 * 9525),
      height: asEmu(120 * 9525),
    });
    await editor.applyTextBodyDocJson(shape.handle, {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: {},
          content: [{ type: "text", text: "Browser edited", marks: [] }],
        },
      ],
    });

    expect(editor.history.undoDepth).toBe(2);
    expect(editor.slides[0]?.svg).toContain("Browser edited");

    expect((await editor.undo()).history).toMatchObject({ canRedo: true, undoDepth: 1 });
    expect(firstText(editor.document)).toBe("Original");
    expect((await editor.redo()).history).toMatchObject({ canUndo: true, redoDepth: 0 });
    expect(firstText(editor.document)).toBe("Browser edited");

    const saved = readPptx(editor.save().pptx);
    expect(firstText(saved)).toBe("Browser edited");
    expect(firstShape(saved).transform).toMatchObject({
      offsetX: 120 * 9525,
      offsetY: 208 * 9525,
      width: 336 * 9525,
      height: 120 * 9525,
    });
  });

  it("returns shared media warnings from image replacement commands", async () => {
    const editor = await createBrowserPptxEditorSession(await buildImageFixture(), {
      skipSystemFonts: true,
    });
    const image = editor.shapes(1).find((shape) => shape.kind === "image");
    if (image?.handle === undefined) throw new Error("image handle not found");
    expect(image.editableImageReplacement).toEqual({
      contentType: "image/png",
      accept: "image/png,.png",
      mediaPartPath: "ppt/media/image1.png",
      sharedReferenceCount: 2,
    });

    const result = await editor.apply({
      kind: "replaceImage",
      handle: image.handle,
      bytes: BLUE_PNG,
    });

    expect(result.warnings).toEqual([
      expect.objectContaining({
        code: "shared-media-part",
        mediaPartPath: "ppt/media/image1.png",
        referenceCount: 2,
      }),
    ]);
    expect(mediaBytes(editor.document, "ppt/media/image1.png")).toEqual(BLUE_PNG);
  });

  it("adds, selects, edits, moves, resizes, saves, deletes, undoes, and redoes a text box", async () => {
    const editor = await createBrowserPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });

    const addedResponse = await editor.addTextBox(1);
    const addedHandle = addedResponse.selection?.shapeHandle;
    if (addedHandle === undefined) throw new Error("added shape was not selected");
    const addedShape = editor
      .shapes(1)
      .find((shape) => handleKey(shape.handle) === handleKey(addedHandle));
    if (addedShape?.handle === undefined) throw new Error("added shape handle not found");
    expect(addedShape).toMatchObject({
      bounds: { x: 96, y: 96, width: 288, height: 72 },
      editableDelete: true,
    });
    expect(addedShape.editableTextBody).toBeDefined();

    await editor.applyTextBodyDocJson(addedShape.handle, {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: {},
          content: [{ type: "text", text: "Added edited", marks: [] }],
        },
      ],
    });
    await editor.apply({
      kind: "setShapeTransform",
      handle: addedShape.handle,
      offsetX: asEmu(144 * 9525),
      offsetY: asEmu(120 * 9525),
      width: asEmu(240 * 9525),
      height: asEmu(96 * 9525),
    });

    const saved = readPptx(editor.save().pptx);
    const savedAdded = shapeByText(saved, "Added edited");
    expect(savedAdded.transform).toMatchObject({
      offsetX: 144 * 9525,
      offsetY: 120 * 9525,
      width: 240 * 9525,
      height: 96 * 9525,
    });

    const deleted = await editor.deleteSelectedShape();
    expect(deleted.selection).toBeUndefined();
    expect(
      editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(addedShape.handle)),
    ).toBe(false);

    await editor.undo();
    expect(
      editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(addedShape.handle)),
    ).toBe(true);
    await editor.redo();
    expect(
      editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(addedShape.handle)),
    ).toBe(false);
  });

  it("adds, selects, moves, resizes, saves, deletes, undoes, and redoes a connector", async () => {
    const editor = await createBrowserPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });

    const addedResponse = await editor.addConnector(1);
    const addedHandle = addedResponse.selection?.shapeHandle;
    if (addedHandle === undefined) throw new Error("added connector was not selected");
    const addedConnector = editor
      .shapes(1)
      .find((shape) => handleKey(shape.handle) === handleKey(addedHandle));
    if (addedConnector?.handle === undefined) throw new Error("added connector handle not found");
    expect(addedConnector).toMatchObject({
      kind: "connector",
      bounds: { x: 144, y: 144, width: 288, height: 96 },
      editableDelete: true,
      editableTransform: true,
    });

    await editor.apply({
      kind: "setShapeTransform",
      handle: addedConnector.handle,
      offsetX: asEmu(168 * 9525),
      offsetY: asEmu(160 * 9525),
      width: asEmu(336 * 9525),
      height: asEmu(120 * 9525),
    });

    const saved = readPptx(editor.save().pptx);
    const savedConnector = connectorByName(saved, "Connector 11");
    expect(savedConnector.transform).toMatchObject({
      offsetX: 168 * 9525,
      offsetY: 160 * 9525,
      width: 336 * 9525,
      height: 120 * 9525,
    });
    expect(savedConnector.outline?.tailEnd).toMatchObject({ type: "triangle" });

    const deleted = await editor.deleteSelectedShape();
    expect(deleted.selection).toBeUndefined();
    expect(editor.shapes(1).find((shape) => shape.name === "Connector 11")).toBeUndefined();

    await editor.undo();
    expect(editor.shapes(1).find((shape) => shape.name === "Connector 11")).toBeDefined();
    await editor.redo();
    expect(editor.shapes(1).find((shape) => shape.name === "Connector 11")).toBeUndefined();
  });

  it("marks top-level text shapes without transform as deletable", async () => {
    const editor = await createBrowserPptxEditorSession(
      await buildShapeFixture({ includeNoTransformShape: true }),
      {
        skipSystemFonts: true,
      },
    );

    const noTransformShape = editor.shapes(1).find((shape) => shape.name === "No Transform");
    if (noTransformShape?.handle === undefined) throw new Error("no-transform shape not found");
    expect(noTransformShape.bounds).toBeUndefined();
    expect(noTransformShape.editableTransform).toBeUndefined();
    expect(noTransformShape.editableDelete).toBe(true);

    await editor.deleteShape(noTransformShape.handle);
    expect(editor.shapes(1).find((shape) => shape.name === "No Transform")).toBeUndefined();
  });

  it("does not mark connector-referenced shapes as deletable", async () => {
    const editor = await createBrowserPptxEditorSession(
      await buildShapeFixture({ includeConnector: true }),
      {
        skipSystemFonts: true,
      },
    );

    const connectedShape = editor.shapes(1).find((shape) => shape.name === "Box");
    if (connectedShape?.handle === undefined) throw new Error("connected shape not found");
    expect(connectedShape.editableDelete).toBeUndefined();
    await expect(editor.deleteShape(connectedShape.handle)).rejects.toThrow(
      /referenced by connector/,
    );
  });

  it("duplicates and deletes slides with render state and history updates", async () => {
    const editor = await createBrowserPptxEditorSession(await buildTwoSlideFixture(), {
      skipSystemFonts: true,
    });
    const firstSlide = editor.slides[0];
    if (firstSlide?.handle === undefined) throw new Error("first slide handle not found");

    const duplicated = await editor.apply({ kind: "duplicateSlide", handle: firstSlide.handle });
    expect(duplicated.slides).toHaveLength(3);
    expect(duplicated.slides.map((slide) => slide.handle?.partPath)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(duplicated.slides[0]?.svg).toContain("First");
    expect(duplicated.slides[1]?.svg).toContain("First");
    expect(duplicated.history).toMatchObject({ canUndo: true, undoDepth: 1 });

    const duplicateSlide = duplicated.slides[1];
    if (duplicateSlide?.handle === undefined) throw new Error("duplicate slide handle not found");
    const deleted = await editor.apply({ kind: "deleteSlide", handle: duplicateSlide.handle });
    expect(deleted.slides.map((slide) => slide.handle?.partPath)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(deleted.history.undoDepth).toBe(2);

    expect((await editor.undo()).slides.map((slide) => slide.handle?.partPath)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect((await editor.redo()).slides.map((slide) => slide.handle?.partPath)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
  });

  it("rejects deleting the last slide without changing the browser editor state", async () => {
    const editor = await createBrowserPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const slide = editor.slides[0];
    if (slide?.handle === undefined) throw new Error("slide handle not found");

    await expect(editor.apply({ kind: "deleteSlide", handle: slide.handle })).rejects.toThrow(
      /last slide/,
    );
    expect(editor.slides).toHaveLength(1);
    expect(editor.history).toMatchObject({ canUndo: false, undoDepth: 0 });
  });

  it("counts unparsed image relationships in image replacement metadata", async () => {
    const editor = await createBrowserPptxEditorSession(
      await buildImageFixture({ includeUnusedImageRelationship: true }),
      { skipSystemFonts: true },
    );
    const image = editor.shapes(1).find((shape) => shape.kind === "image");

    expect(image?.editableImageReplacement?.sharedReferenceCount).toBe(3);
  });
});

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function pngBytes(base64: string): Uint8Array {
  return new Uint8Array(Buffer.from(base64, "base64"));
}

async function buildShapeFixture(
  options: { includeNoTransformShape?: boolean; includeConnector?: boolean } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:r><a:rPr sz="2400"><a:latin typeface="Aptos"/></a:rPr><a:t>Original</a:t></a:r></a:p>` +
        `</p:txBody>` +
        `</p:sp>` +
        (options.includeNoTransformShape
          ? `<p:sp><p:nvSpPr><p:cNvPr id="11" name="No Transform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
            `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
            `<p:txBody><a:bodyPr/><a:lstStyle/>` +
            `<a:p><a:r><a:t>No transform</a:t></a:r></a:p>` +
            `</p:txBody>` +
            `</p:sp>`
          : "") +
        (options.includeConnector
          ? `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="12" name="Connector"/><p:cNvCxnSpPr>` +
            `<a:stCxn id="10" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
            `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
            `<a:prstGeom prst="straightConnector1"/></p:spPr></p:cxnSp>`
          : "") +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

async function buildImageFixture(
  options: { readonly includeUnusedImageRelationship?: boolean } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Shared Picture A"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="21" name="Shared Picture B"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        (options.includeUnusedImageRelationship === true
          ? `<Relationship Id="rIdUnusedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`
          : "") +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/media/image1.png", RED_PNG);

  return zip.generateAsync({ type: "uint8array" });
}

async function buildTwoSlideFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/slides/slide1.xml", textSlideXml(10, "First"));
  zip.file("ppt/slides/slide2.xml", textSlideXml(20, "Second"));

  return zip.generateAsync({ type: "uint8array" });
}

function textSlideXml(shapeId: number, text: string): Uint8Array {
  return xml(
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<p:cSld><p:spTree>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="${String(shapeId)}" name="${text}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
      `</p:sp>` +
      `</p:spTree></p:cSld>` +
      `</p:sld>`,
  );
}

function firstShape(source: ReturnType<typeof readPptx>): SourceShape {
  const shape = source.slides[0]?.shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("fixture shape not found");
  return shape;
}

function firstText(source: ReturnType<typeof readPptx>): string {
  const run = firstShape(source).textBody?.paragraphs[0]?.runs[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.text;
}

function shapeByText(source: ReturnType<typeof readPptx>, text: string): SourceShape {
  const shape = source.slides[0]?.shapes.find(
    (node): node is SourceShape =>
      node.kind === "shape" &&
      node.textBody?.paragraphs.some((paragraph) =>
        paragraph.runs.some((run) => run.text === text),
      ) === true,
  );
  if (shape === undefined) throw new Error(`shape text not found: ${text}`);
  return shape;
}

function connectorByName(source: ReturnType<typeof readPptx>, name: string): SourceConnector {
  const connector = source.slides[0]?.shapes.find(
    (node): node is SourceConnector => node.kind === "connector" && node.name === name,
  );
  if (connector === undefined) throw new Error(`connector not found: ${name}`);
  return connector;
}

function handleKey(handle: unknown): string {
  if (handle === undefined || handle === null || typeof handle !== "object") return "";
  const value = handle as {
    partPath?: string;
    nodeId?: string;
    relationshipId?: string;
    orderingSlot?: number;
  };
  return [
    value.partPath ?? "",
    value.nodeId ?? "",
    value.relationshipId ?? "",
    value.orderingSlot ?? "",
  ].join("\u0000");
}

function mediaBytes(source: ReturnType<typeof readPptx>, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}
