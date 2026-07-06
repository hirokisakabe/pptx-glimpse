import { describe, expect, it } from "vitest";

import {
  addConnector,
  addEmptySlideFromLayout,
  addTextBox,
  clearTextRunProperties,
  deleteShape,
  deleteSlide,
  duplicateSlide,
  findParagraphBySourceHandle,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  moveSlide,
  replaceImageBytes,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setTextRunProperties,
  updateShapeTransform,
} from "./editing.js";
import { asPartPath, asRelationshipId, asSourceNodeId, type SourceHandle } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceImage, SourceShape } from "./shapes.js";
import { asEmu, asPt } from "./units.js";

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const PRESENTATION_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const SLIDE_LAYOUT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 1, 2, 3]);

describe("editing source helpers", () => {
  it("finds text runs, paragraphs, and shape nodes by source handle", () => {
    const source = buildSourceModel();
    const run = firstRun(source);
    const paragraph = firstParagraph(source);
    const shape = shapeByName(source, "Title");

    expect(findTextRunBySourceHandle(source, requireHandle(run.handle))).toBe(run);
    expect(findParagraphBySourceHandle(source, requireHandle(paragraph.handle))).toBe(paragraph);
    expect(findShapeNodeBySourceHandle(source, requireHandle(shape.handle))).toBe(shape);
    expect(findTextRunBySourceHandle(source, handle("ppt/slides/slide1.xml", "missing"))).toBe(
      undefined,
    );
  });
});

describe("editing text operations", () => {
  it("records a text run replacement without mutating the input source", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const edited = expectNonMutating(source, () =>
      replaceTextRunPlainText(source, runHandle, "Edited"),
    );

    expect(firstRun(source).text).toBe("Original");
    expect(firstRun(edited).text).toBe("Edited");
    expect(edited.edits).toEqual([
      { kind: "replaceTextRunPlainText", handle: runHandle, text: "Edited" },
    ]);
  });

  it("returns the original source when replacing a text run with the same text", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const edited = replaceTextRunPlainText(source, runHandle, firstRun(source).text);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("does not append another text run edit when the replacement is already applied", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);
    const edited = replaceTextRunPlainText(source, runHandle, "Edited");

    const repeated = replaceTextRunPlainText(edited, runHandle, "Edited");

    expect(repeated).toBe(edited);
    expect(repeated.edits).toHaveLength(1);
  });

  it("sets and clears text run properties in source and edit records", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const set = expectNonMutating(source, () =>
      setTextRunProperties(source, runHandle, {
        bold: false,
        underline: true,
        fontSize: asPt(18),
        color: { kind: "srgb", hex: "00aa44" },
        typeface: "Aptos",
      }),
    );
    const cleared = expectNonMutating(set, () =>
      clearTextRunProperties(set, runHandle, ["italic", "typeface"]),
    );

    expect(firstRun(set).properties).toMatchObject({
      bold: false,
      italic: true,
      underline: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Aptos",
    });
    expect(firstRun(cleared).properties).toMatchObject({
      bold: false,
      underline: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "00aa44" },
    });
    expect(firstRun(cleared).properties?.italic).toBeUndefined();
    expect(firstRun(cleared).properties?.typeface).toBeUndefined();
    expect(cleared.edits?.map((edit) => edit.kind)).toEqual([
      "updateTextRunProperties",
      "updateTextRunProperties",
    ]);
    expect(cleared.edits?.[0]).toMatchObject({
      kind: "updateTextRunProperties",
      set: {
        bold: false,
        underline: true,
        fontSize: 18,
        color: { kind: "srgb", hex: "00aa44" },
        typeface: "Aptos",
      },
    });
    expect(cleared.edits?.[1]).toMatchObject({
      kind: "updateTextRunProperties",
      clear: ["italic", "typeface"],
    });
  });

  it("returns the original source when clearing properties that are already absent", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const edited = clearTextRunProperties(source, runHandle, ["underline"]);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("returns the original source for equivalent text properties with different key order", () => {
    const source = structuredClone(buildSourceModel());
    const runHandle = requireHandle(firstRun(source).handle);
    (firstRun(source) as { properties?: unknown }).properties = {
      color: { hex: "00aa44", kind: "srgb" },
      bold: true,
      italic: true,
      fontSize: asPt(24),
      typeface: "Calibri",
    };

    const edited = setTextRunProperties(source, runHandle, {
      color: { kind: "srgb", hex: "00aa44" },
    });

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("supports text edits addressed by shapeSlot handles", () => {
    const source = buildSourceModel();
    const slotShape = shapeByName(source, "Slot Text");
    const slotRun = slotShape.textBody?.paragraphs[0]?.runs[0];
    const slotRunHandle = requireHandle(slotRun?.handle);

    expect(slotRunHandle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shapeSlot:3:p:0:r:0",
      orderingSlot: 0,
    });
    expect(findTextRunBySourceHandle(source, slotRunHandle)).toBe(slotRun);

    const edited = expectNonMutating(source, () =>
      replaceTextRunPlainText(source, slotRunHandle, "Slot Edited"),
    );

    expect(shapeByName(edited, "Slot Text").textBody?.paragraphs[0]?.runs[0]?.text).toBe(
      "Slot Edited",
    );
    expect(edited.edits).toEqual([
      { kind: "replaceTextRunPlainText", handle: slotRunHandle, text: "Slot Edited" },
    ]);
  });

  it("normalizes a paragraph replacement to one run and records the edit", () => {
    const source = buildSourceModel();
    const paragraphHandle = requireHandle(firstParagraph(source).handle);

    const edited = expectNonMutating(source, () =>
      replaceParagraphPlainText(source, paragraphHandle, "Paragraph"),
    );

    expect(firstParagraph(source).runs).toHaveLength(2);
    expect(firstParagraph(edited).runs).toEqual([
      {
        kind: "textRun",
        text: "Paragraph",
        properties: firstRun(source).properties,
        handle: firstRun(source).handle,
      },
    ]);
    expect(edited.edits).toEqual([
      { kind: "replaceParagraphPlainText", handle: paragraphHandle, text: "Paragraph" },
    ]);
  });

  it("rejects invalid text property patches with operation-specific errors", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    expect(() => setTextRunProperties(source, runHandle, { fontSize: asPt(0) })).toThrow(
      "updateTextRunProperties: fontSize must be a finite positive pt value",
    );
    expect(() =>
      setTextRunProperties(source, runHandle, { color: { kind: "srgb", hex: "abc" } }),
    ).toThrow("updateTextRunProperties: srgb text run color must be a 6-digit hex value");
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      clearTextRunProperties(source, runHandle, ["strikethrough"]),
    ).toThrow("updateTextRunProperties: unsupported text run property 'strikethrough'");
  });
});

describe("editing shape operations", () => {
  it("updates a top-level shape transform and records the edit", () => {
    const source = buildSourceModel();
    const shapeHandle = requireHandle(shapeByName(source, "Title").handle);
    const transform = {
      offsetX: asEmu(11),
      offsetY: asEmu(22),
      width: asEmu(33),
      height: asEmu(44),
    };

    const edited = expectNonMutating(source, () =>
      updateShapeTransform(source, shapeHandle, transform),
    );

    expect(shapeByName(source, "Title").transform).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 3000,
      height: 1000,
    });
    expect(shapeByName(edited, "Title").transform).toMatchObject(transform);
    expect(edited.edits).toEqual([
      { kind: "updateShapeTransform", handle: shapeHandle, ...transform },
    ]);
  });

  it("returns the original source when updating a shape transform to the same values", () => {
    const source = buildSourceModel();
    const shape = shapeByName(source, "Title");
    const shapeHandle = requireHandle(shape.handle);
    const transform = {
      offsetX: shape.transform!.offsetX,
      offsetY: shape.transform!.offsetY,
      width: shape.transform!.width,
      height: shape.transform!.height,
    };

    const edited = updateShapeTransform(source, shapeHandle, transform);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("does not append another shape transform edit when the transform is already applied", () => {
    const source = buildSourceModel();
    const shapeHandle = requireHandle(shapeByName(source, "Title").handle);
    const transform = {
      offsetX: asEmu(11),
      offsetY: asEmu(22),
      width: asEmu(33),
      height: asEmu(44),
    };
    const edited = updateShapeTransform(source, shapeHandle, transform);

    const repeated = updateShapeTransform(edited, shapeHandle, transform);

    expect(repeated).toBe(edited);
    expect(repeated.edits).toHaveLength(1);
  });

  it("adds a text box with a collision-free id and finalized XML", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addTextBox(source, requireHandle(source.slides[0].handle), {
        offsetX: asEmu(500),
        offsetY: asEmu(600),
        width: asEmu(700),
        height: asEmu(800),
        text: "007",
      }),
    );
    const added = shapeByName(edited, "TextBox 31");

    expect(
      source.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("TextBox 31");
    expect(added.nodeId).toBe("31");
    expect(added.textBody?.paragraphs[0]?.runs[0]?.text).toBe("007");
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "addTextBox",
      slidePartPath: "ppt/slides/slide1.xml",
      shapeId: "31",
    });
    expect(edited.edits?.at(-1)).toHaveProperty("xml", expect.stringContaining("TextBox 31"));
  });

  it("adds a connector using source target handles and records finalized XML", () => {
    const source = buildSourceModel();
    const start = shapeByName(source, "Title");
    const end = shapeByName(source, "Body");

    const edited = expectNonMutating(source, () =>
      addConnector(source, requireHandle(source.slides[0].handle), {
        preset: "bentConnector3",
        offsetX: asEmu(10),
        offsetY: asEmu(20),
        width: asEmu(30),
        height: asEmu(40),
        start: { shapeHandle: requireHandle(start.handle), connectionSiteIndex: 1 },
        end: { shapeHandle: requireHandle(end.handle), connectionSiteIndex: 3 },
        outline: { tailEnd: { type: "triangle", width: "med", length: "lg" } },
      }),
    );
    const connector = edited.slides[0].shapes.at(-1);

    expect(connector).toMatchObject({
      kind: "connector",
      nodeId: "31",
      name: "Connector 31",
      connection: {
        start: { shapeId: "10", connectionSiteIndex: 1 },
        end: { shapeId: "30", connectionSiteIndex: 3 },
      },
      geometry: { preset: "bentConnector3" },
    });
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "addConnector",
      slidePartPath: "ppt/slides/slide1.xml",
      shapeId: "31",
      startShapeId: "10",
      endShapeId: "30",
    });
    expect(edited.edits?.at(-1)).toHaveProperty("xml", expect.stringContaining("bentConnector3"));
  });

  it("adds a free connector without native connection sites", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addConnector(source, requireHandle(source.slides[0].handle), {
        preset: "straightConnector1",
        offsetX: asEmu(10),
        offsetY: asEmu(20),
        width: asEmu(30),
        height: asEmu(40),
        outline: { tailEnd: { type: "triangle", width: "med", length: "med" } },
      }),
    );
    const connector = edited.slides[0].shapes.at(-1);

    expect(connector).toMatchObject({
      kind: "connector",
      nodeId: "31",
      name: "Connector 31",
      geometry: { preset: "straightConnector1" },
      outline: { tailEnd: { type: "triangle" } },
    });
    expect(connector).not.toHaveProperty("connection.start");
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "addConnector",
      slidePartPath: "ppt/slides/slide1.xml",
      shapeId: "31",
    });
    expect(edited.edits?.at(-1)).not.toHaveProperty("startShapeId");
    expect(edited.edits?.at(-1)).not.toHaveProperty("endShapeId");
  });

  it("deletes an existing shape, drops stale text edits, and cancels added shapes followed by deleteShape", () => {
    const source = buildSourceModel();
    const existingHandle = requireHandle(shapeByName(source, "Title").handle);

    const deletedExisting = expectNonMutating(source, () => deleteShape(source, existingHandle));
    const withTextEdit = replaceTextRunPlainText(
      source,
      requireHandle(firstRun(source).handle),
      "Stale",
    );
    const deletedTextEditedShape = expectNonMutating(withTextEdit, () =>
      deleteShape(withTextEdit, existingHandle),
    );
    const withTextBox = addTextBox(source, requireHandle(source.slides[0].handle), {
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      text: "Temporary",
      name: "Temporary",
    });
    const deletedAdded = expectNonMutating(withTextBox, () =>
      deleteShape(withTextBox, requireHandle(shapeByName(withTextBox, "Temporary").handle)),
    );
    const withConnector = addConnector(source, requireHandle(source.slides[0].handle), {
      preset: "straightConnector1",
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      name: "Temporary connector",
    });
    const deletedAddedConnector = expectNonMutating(withConnector, () =>
      deleteShape(
        withConnector,
        requireHandle(
          withConnector.slides[0].shapes.find(
            (shape) => shape.kind === "connector" && shape.name === "Temporary connector",
          )?.handle,
        ),
      ),
    );

    expect(
      deletedExisting.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Title");
    expect(deletedExisting.edits).toEqual([{ kind: "deleteShape", handle: existingHandle }]);
    expect(deletedTextEditedShape.edits).toEqual([{ kind: "deleteShape", handle: existingHandle }]);
    expect(deletedAdded.edits).toEqual([]);
    expect(
      deletedAdded.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Temporary");
    expect(deletedAddedConnector.edits).toEqual([]);
    expect(
      deletedAddedConnector.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Temporary connector");
  });

  it("rejects unsupported shape edits with stable error messages", () => {
    const source = buildSourceModel();
    const imageHandle = requireHandle(imageByName(source, "Picture").handle);
    const noNodeHandle = { partPath: asPartPath("ppt/slides/slide1.xml") };

    expect(() =>
      updateShapeTransform(source, noNodeHandle, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow("updateShapeTransform: shape transform edit requires a node id");
    expect(() => deleteShape(source, imageHandle)).toThrow(
      "deleteShape: only top-level sp or cxnSp shapes can be deleted",
    );
    expect(() =>
      addConnector(source, requireHandle(source.slides[0].handle), {
        // @ts-expect-error exercises runtime validation for JS callers.
        preset: "arc",
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        start: {
          shapeHandle: requireHandle(shapeByName(source, "Title").handle),
          connectionSiteIndex: 0,
        },
        end: {
          shapeHandle: requireHandle(shapeByName(source, "Body").handle),
          connectionSiteIndex: 0,
        },
      }),
    ).toThrow(
      "addConnector: preset must be straightConnector1, bentConnector3, or curvedConnector3",
    );
  });
});

describe("editing media and slide topology operations", () => {
  it("replaces image bytes after validating content type and records shared reference count", () => {
    const source = buildSourceModel();
    const imageHandle = requireHandle(imageByName(source, "Picture").handle);

    const edited = expectNonMutating(source, () => replaceImageBytes(source, imageHandle, PNG));

    expect(source.packageGraph.media[0].bytes).not.toBe(PNG);
    expect(edited.packageGraph.media[0]).toMatchObject({
      partPath: "ppt/media/image1.png",
      contentType: "image/png",
      bytes: PNG,
    });
    expect(edited.packageGraph.media[0].bytes).not.toBe(PNG);
    expect(edited.edits).toEqual([
      {
        kind: "replaceImage",
        handle: imageHandle,
        mediaPartPath: "ppt/media/image1.png",
        contentType: "image/png",
        sharedReferenceCount: 1,
      },
    ]);
    expect(() => replaceImageBytes(source, imageHandle, JPEG)).toThrow(
      "replaceImageBytes: replacement image content type 'image/jpeg' does not match existing media content type 'image/png'",
    );
  });

  it("adds an empty slide from a layout with reserved part, relationship, and numeric ids", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addEmptySlideFromLayout(source, {
        layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout1.xml"),
      }),
    );

    expect(source.slides).toHaveLength(2);
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
    ]);
    expect(edited.slides[2]).toMatchObject({
      partPath: "ppt/slides/slide3.xml",
      layoutPartPath: "ppt/slideLayouts/slideLayout1.xml",
      shapes: [],
      handle: { partPath: "ppt/slides/slide3.xml" },
    });
    expect(edited.edits?.at(-1)).toEqual({
      kind: "addEmptySlideFromLayout",
      layoutPartPath: "ppt/slideLayouts/slideLayout1.xml",
      newSlidePartPath: "ppt/slides/slide3.xml",
      newRelationshipId: "rId3",
      newSlideNumericId: 301,
    });
  });

  it("duplicates a slide after the source slide and rewrites cloned handles to the new part", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      duplicateSlide(source, requireHandle(source.slides[0].handle)),
    );

    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.slides[1].partPath).toBe("ppt/slides/slide3.xml");
    expect(edited.slides[1].shapes[0]?.handle?.partPath).toBe("ppt/slides/slide3.xml");
    const clonedShape = edited.slides[1].shapes[0];
    if (clonedShape?.kind !== "shape") throw new Error("expected duplicated text shape");
    expect(clonedShape.textBody?.paragraphs[0]?.handle?.partPath).toBe("ppt/slides/slide3.xml");
    expect(clonedShape.textBody?.paragraphs[0]?.runs[0]?.handle?.partPath).toBe(
      "ppt/slides/slide3.xml",
    );
    expect(edited.edits?.at(-1)).toEqual({
      kind: "duplicateSlide",
      sourceSlidePartPath: "ppt/slides/slide1.xml",
      sourceRelationshipId: "rIdSlide1",
      newSlidePartPath: "ppt/slides/slide3.xml",
      newRelationshipId: "rId3",
      newSlideNumericId: 301,
    });
  });

  it("moves a slide to the requested final index without changing package parts", () => {
    const source = buildSourceModel();
    const edited = expectNonMutating(source, () =>
      moveSlide(source, requireHandle(source.slides[0].handle), { toIndex: 1 }),
    );

    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(edited.slides.map((slide) => slide.partPath)).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(edited.packageGraph.parts).toEqual(source.packageGraph.parts);
    expect(edited.edits?.at(-1)).toEqual({
      kind: "moveSlide",
      slidePartPath: "ppt/slides/slide1.xml",
      relationshipId: "rIdSlide1",
      toIndex: 1,
    });
  });

  it("treats moving a slide to its current index as a no-op", () => {
    const source = buildSourceModel();
    const edited = moveSlide(source, requireHandle(source.slides[1].handle), { toIndex: 1 });

    expect(edited).toBe(source);
    expect(edited.edits).toBeUndefined();
  });

  it("deletes a slide while dropping pending edits invalidated by that part", () => {
    const source = buildSourceModel();
    const dirty = replaceTextRunPlainText(source, requireHandle(firstRun(source).handle), "Dirty");

    const edited = expectNonMutating(dirty, () =>
      deleteSlide(dirty, requireHandle(dirty.slides[0].handle)),
    );

    expect(dirty.edits?.map((edit) => edit.kind)).toEqual(["replaceTextRunPlainText"]);
    expect(edited.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
    expect(edited.edits).toEqual([
      {
        kind: "deleteSlide",
        slidePartPath: "ppt/slides/slide1.xml",
        relationshipId: "rIdSlide1",
      },
    ]);
    expect(edited.packageGraph.parts.map((part) => part.partPath)).not.toContain(
      "ppt/slides/slide1.xml",
    );
  });

  it("cancels addEmptySlideFromLayout followed by deleteSlide", () => {
    const source = buildSourceModel();
    const added = addEmptySlideFromLayout(source, {
      layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout1.xml"),
    });
    const edited = expectNonMutating(added, () =>
      deleteSlide(added, requireHandle(added.slides[2].handle)),
    );

    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.edits).toEqual([]);
    expect(edited.packageGraph.parts.map((part) => part.partPath)).not.toContain(
      "ppt/slides/slide3.xml",
    );
  });

  it("rejects invalid slide topology operations with operation-specific errors", () => {
    const source = buildSourceModel();
    const singleSlide = { ...source, slides: [source.slides[0]] };

    expect(() =>
      addEmptySlideFromLayout(source, {
        layoutPartPath: asPartPath("ppt/slideLayouts/missing.xml"),
      }),
    ).toThrow("addEmptySlideFromLayout: slide layout part path was not found");
    expect(() =>
      duplicateSlide(
        replaceTextRunPlainText(source, requireHandle(firstRun(source).handle), "Dirty"),
        requireHandle(source.slides[0].handle),
      ),
    ).toThrow("duplicateSlide: duplicating a slide with pending dirty part edits is unsupported");
    expect(() => deleteSlide(singleSlide, requireHandle(singleSlide.slides[0].handle))).toThrow(
      "deleteSlide: cannot delete the last slide",
    );
    expect(() =>
      moveSlide(source, { partPath: asPartPath("ppt/slides/missing.xml") }, { toIndex: 1 }),
    ).toThrow("moveSlide: slide handle was not found in PptxSourceModel source");
    expect(() => moveSlide(source, requireHandle(source.slides[0].handle), { toIndex: 2 })).toThrow(
      "moveSlide: toIndex must be an integer slide index in range",
    );
  });
});

function expectNonMutating(source: PptxSourceModel, edit: () => PptxSourceModel): PptxSourceModel {
  const before = structuredClone(source);
  const edited = edit();

  expect(source).toEqual(before);
  if (edited !== source) expect(edited).not.toBe(source);
  return edited;
}

function buildSourceModel(): PptxSourceModel {
  const presentationPath = asPartPath("ppt/presentation.xml");
  const slide1Path = asPartPath("ppt/slides/slide1.xml");
  const slide2Path = asPartPath("ppt/slides/slide2.xml");
  const layoutPath = asPartPath("ppt/slideLayouts/slideLayout1.xml");
  const imagePath = asPartPath("ppt/media/image1.png");

  return {
    packageGraph: {
      contentTypes: {
        defaults: [
          {
            extension: "rels",
            contentType: "application/vnd.openxmlformats-package.relationships+xml",
          },
          { extension: "png", contentType: "image/png" },
        ],
        overrides: [
          { partName: presentationPath, contentType: PRESENTATION_CONTENT_TYPE },
          { partName: slide1Path, contentType: SLIDE_CONTENT_TYPE },
          { partName: slide2Path, contentType: SLIDE_CONTENT_TYPE },
          { partName: layoutPath, contentType: SLIDE_LAYOUT_CONTENT_TYPE },
        ],
      },
      parts: [
        { partPath: presentationPath, contentType: PRESENTATION_CONTENT_TYPE },
        { partPath: slide1Path, contentType: SLIDE_CONTENT_TYPE },
        { partPath: slide2Path, contentType: SLIDE_CONTENT_TYPE },
        { partPath: layoutPath, contentType: SLIDE_LAYOUT_CONTENT_TYPE },
        { partPath: imagePath, contentType: "image/png" },
      ],
      relationships: [
        {
          sourcePartPath: presentationPath,
          relationships: [
            {
              id: asRelationshipId("rIdSlide1"),
              type: SLIDE_REL_TYPE,
              target: "slides/slide1.xml",
            },
            {
              id: asRelationshipId("rIdSlide2"),
              type: SLIDE_REL_TYPE,
              target: "slides/slide2.xml",
            },
          ],
        },
        {
          sourcePartPath: slide1Path,
          relationships: [
            {
              id: asRelationshipId("rIdLayout"),
              type: SLIDE_LAYOUT_REL_TYPE,
              target: "../slideLayouts/slideLayout1.xml",
            },
            {
              id: asRelationshipId("rIdImage"),
              type: IMAGE_REL_TYPE,
              target: "../media/image1.png",
            },
          ],
        },
        {
          sourcePartPath: slide2Path,
          relationships: [
            {
              id: asRelationshipId("rIdLayout"),
              type: SLIDE_LAYOUT_REL_TYPE,
              target: "../slideLayouts/slideLayout1.xml",
            },
          ],
        },
      ],
      media: [{ partPath: imagePath, contentType: "image/png", bytes: new Uint8Array(PNG) }],
      rawParts: [
        {
          kind: "binary",
          partPath: presentationPath,
          contentType: PRESENTATION_CONTENT_TYPE,
          bytes: xml(
            `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
              `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
              `<p:sldIdLst>` +
              `<p:sldId id="256" r:id="rIdSlide1"/>` +
              `<p:sldId id="300" r:id="rIdSlide2"/>` +
              `</p:sldIdLst>` +
              `</p:presentation>`,
          ),
        },
        {
          kind: "binary",
          partPath: slide1Path,
          contentType: SLIDE_CONTENT_TYPE,
          bytes: xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
        },
        {
          kind: "binary",
          partPath: slide2Path,
          contentType: SLIDE_CONTENT_TYPE,
          bytes: xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
        },
      ],
    },
    presentation: {
      partPath: presentationPath,
      slideSize: { width: asEmu(9144000), height: asEmu(5143500) },
      slidePartPaths: [slide1Path, slide2Path],
    },
    slides: [
      {
        partPath: slide1Path,
        layoutPartPath: layoutPath,
        shapes: [
          textShape(slide1Path, "10", "Title", 0),
          imageShape(slide1Path),
          textShape(slide1Path, "30", "Body", 2),
          slotTextShape(slide1Path, 3),
        ],
        handle: { partPath: slide1Path },
      },
      {
        partPath: slide2Path,
        layoutPartPath: layoutPath,
        shapes: [textShape(slide2Path, "40", "Second", 0)],
        handle: { partPath: slide2Path },
      },
    ],
    slideLayouts: [
      {
        partPath: layoutPath,
        masterPartPath: asPartPath("ppt/slideMasters/slideMaster1.xml"),
        type: "blank",
        shapes: [],
      },
    ],
    slideMasters: [],
    themes: [],
    diagnostics: [],
  };
}

function textShape(
  partPath: ReturnType<typeof asPartPath>,
  id: string,
  name: string,
  orderingSlot: number,
): SourceShape {
  const nodeId = asSourceNodeId(id);
  return {
    kind: "shape",
    nodeId,
    name,
    transform: {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(3000),
      height: asEmu(1000),
    },
    geometry: { preset: "rect" },
    textBody: {
      paragraphs: [
        {
          handle: handle(partPath, `text:shape:${id}:p:0`, 0),
          runs: [
            {
              kind: "textRun",
              text: "Original",
              properties: { bold: true, italic: true, fontSize: asPt(24), typeface: "Calibri" },
              handle: handle(partPath, `text:shape:${id}:p:0:r:0`, 0),
            },
            {
              kind: "textRun",
              text: " Keep",
              handle: handle(partPath, `text:shape:${id}:p:0:r:1`, 1),
            },
          ],
        },
      ],
    },
    handle: handle(partPath, id, orderingSlot),
  };
}

function slotTextShape(partPath: ReturnType<typeof asPartPath>, orderingSlot: number): SourceShape {
  return {
    kind: "shape",
    name: "Slot Text",
    transform: {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(3000),
      height: asEmu(1000),
    },
    geometry: { preset: "rect" },
    textBody: {
      paragraphs: [
        {
          handle: handle(partPath, `text:shapeSlot:${orderingSlot}:p:0`, 0),
          runs: [
            {
              kind: "textRun",
              text: "Slot Original",
              handle: handle(partPath, `text:shapeSlot:${orderingSlot}:p:0:r:0`, 0),
            },
          ],
        },
      ],
    },
    handle: { partPath, orderingSlot },
  };
}

function imageShape(partPath: ReturnType<typeof asPartPath>): SourceImage {
  return {
    kind: "image",
    nodeId: asSourceNodeId("20"),
    name: "Picture",
    blipRelationshipId: asRelationshipId("rIdImage"),
    handle: {
      partPath,
      nodeId: asSourceNodeId("20"),
      relationshipId: asRelationshipId("rIdImage"),
      orderingSlot: 1,
    },
  };
}

function handle(
  partPath: string | ReturnType<typeof asPartPath>,
  nodeId: string,
  orderingSlot?: number,
): SourceHandle {
  return {
    partPath: typeof partPath === "string" ? asPartPath(partPath) : partPath,
    nodeId: asSourceNodeId(nodeId),
    ...(orderingSlot === undefined ? {} : { orderingSlot }),
  };
}

function xml(content: string): Uint8Array {
  return new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`,
  );
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("test fixture handle is missing");
  return handle;
}

function firstRun(source: PptxSourceModel) {
  return firstParagraph(source).runs[0];
}

function firstParagraph(source: PptxSourceModel) {
  const shape = shapeByName(source, "Title");
  const paragraph = shape.textBody?.paragraphs[0];
  if (paragraph === undefined) throw new Error("test fixture paragraph is missing");
  return paragraph;
}

function shapeByName(source: PptxSourceModel, name: string): SourceShape {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "shape" && shape.name === name) return shape;
    }
  }
  throw new Error(`test fixture shape '${name}' is missing`);
}

function imageByName(source: PptxSourceModel, name: string): SourceImage {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "image" && shape.name === name) return shape;
    }
  }
  throw new Error(`test fixture image '${name}' is missing`);
}
