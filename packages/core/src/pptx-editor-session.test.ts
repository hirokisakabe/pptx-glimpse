import { Buffer } from "node:buffer";

import {
  addEmptySlideFromLayout,
  addShape,
  addSlideLayout,
  asEmu,
  asOoxmlPercent,
  asSourceNodeId,
  createPptx,
  readPptx,
  setBackground,
  setShapeFill,
  type SourceConnector,
  type SourceShape,
  writePptx,
} from "@pptx-glimpse/document";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";
import { createPptxEditorSession, isPptxEditorError, PptxEditorError } from "./index.js";
import {
  affectedSlidePartPaths,
  configurePptxEditorSessionAffectedSlidesResolver,
  configurePptxEditorSessionRenderer,
  type PptxEditorErrorCode,
} from "./pptx-editor-session.js";

const nodeFontMocks = vi.hoisted(() => ({
  createOpentypeSetupFromSystem: vi.fn().mockResolvedValue(null),
}));

vi.mock("@pptx-glimpse/renderer/node", () => ({
  createOpentypeSetupFromSystem: nodeFontMocks.createOpentypeSetupFromSystem,
}));

const encoder = new TextEncoder();
const RED_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=",
);
const BLUE_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
);

describe("PptxEditorSession", () => {
  it("edits, rerenders, and saves Uint8Array PPTX bytes in Node.js", async () => {
    const editor = await createPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const shape = editor.shapes(1)[0];
    if (shape?.handle === undefined) throw new Error("shape handle not found");

    expect(editor.slides).toHaveLength(1);
    expect(shape.bounds).toEqual({ x: 96, y: 192, width: 288, height: 96 });
    expect(shape.textBody).toMatchObject({
      paragraphs: [
        {
          runs: [
            {
              text: "Original",
              properties: {
                fontSize: 24,
                typeface: "Aptos",
              },
            },
          ],
        },
      ],
    });
    expect(shape.textBody?.paragraphs[0]?.handle).toBeDefined();
    expect(shape.textBody?.paragraphs[0]?.runs[0]?.handle).toBeDefined();
    const runHandle = shape.textBody?.paragraphs[0]?.runs[0]?.handle;
    if (runHandle === undefined) throw new Error("text run handle not found");

    await editor.apply({
      kind: "setShapeTransform",
      handle: shape.handle,
      offsetX: asEmu(120 * 9525),
      offsetY: asEmu(208 * 9525),
      width: asEmu(336 * 9525),
      height: asEmu(120 * 9525),
    });
    await editor.apply({
      kind: "replaceTextRunPlainText",
      handle: runHandle,
      text: "Node edited",
    });

    expect(editor.history.undoDepth).toBe(2);
    expect(editor.slides[0]?.svg).toContain("Node edited");

    expect((await editor.undo()).history).toMatchObject({ canRedo: true, undoDepth: 1 });
    expect(firstText(editor.document)).toBe("Original");
    expect((await editor.redo()).history).toMatchObject({ canUndo: true, redoDepth: 0 });
    expect(firstText(editor.document)).toBe("Node edited");

    const saved = readPptx(editor.save().pptx);
    expect(firstText(saved)).toBe("Node edited");
    expect(firstShape(saved).transform).toMatchObject({
      offsetX: 120 * 9525,
      offsetY: 208 * 9525,
      width: 336 * 9525,
      height: 120 * 9525,
    });
  });

  it("uses the Node font loader selected by the main entry", async () => {
    nodeFontMocks.createOpentypeSetupFromSystem.mockClear();

    await createPptxEditorSession(await buildShapeFixture(), {
      fontDirs: ["/app/fonts"],
      skipSystemFonts: true,
    });

    expect(nodeFontMocks.createOpentypeSetupFromSystem).toHaveBeenCalledWith(
      ["/app/fonts"],
      undefined,
      true,
    );
  });

  it("applies command batches as one history entry and renders once", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildShapeFixture(), {
        skipSystemFonts: true,
      });
      const run = editor.shapes(1)[0]?.textBody?.paragraphs[0]?.runs[0];
      if (run?.handle === undefined) throw new Error("text run handle not found");

      const response = await editor.applyAll([
        {
          kind: "replaceTextRunPlainText",
          handle: run.handle,
          text: "Batch edited",
        },
        {
          kind: "setTextRunProperties",
          handle: run.handle,
          properties: { bold: true },
        },
      ]);

      expect(renderCalls).toEqual([undefined, [1]]);
      expect(response.history).toMatchObject({ undoDepth: 1, canUndo: true });
      expect(firstText(editor.document)).toBe("Batch edited");
      expect(firstShape(editor.document).textBody?.paragraphs[0]?.runs[0]?.properties?.bold).toBe(
        true,
      );

      await editor.undo();
      expect(firstText(editor.document)).toBe("Original");
      expect(
        firstShape(editor.document).textBody?.paragraphs[0]?.runs[0]?.properties?.bold,
      ).toBeUndefined();
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("rerenders only the affected slides for applyAll, undo, and redo", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    let renderGeneration = 0;
    configurePptxEditorSessionRenderer(async (source, options) => {
      renderGeneration += 1;
      renderCalls.push(options?.slides);
      const report = await renderPptxSourceModelToSvg(source, options);
      return {
        ...report,
        slides: report.slides.map((slide) => ({
          ...slide,
          svg: `${slide.svg}<!-- render:${String(renderGeneration)} -->`,
        })),
      };
    });
    try {
      const editor = await createPptxEditorSession(await buildTwoSlideFixture(), {
        skipSystemFonts: true,
      });
      const initialFirstSvg = editor.slides[0]?.svg;
      const firstRun = editor.shapes(1)[0]?.textBody?.paragraphs[0]?.runs[0];
      const secondRun = editor.shapes(2)[0]?.textBody?.paragraphs[0]?.runs[0];
      if (firstRun?.handle === undefined || secondRun?.handle === undefined) {
        throw new Error("text run handles not found");
      }

      await editor.apply({
        kind: "replaceTextRunPlainText",
        handle: secondRun.handle,
        text: "Second edited",
      });
      expect(renderCalls).toEqual([undefined, [2]]);
      expect(editor.slides[0]?.svg).toBe(initialFirstSvg);
      expect(editor.slides[1]?.svg).toContain("Second edited");

      await editor.applyAll([
        {
          kind: "replaceTextRunPlainText",
          handle: firstRun.handle,
          text: "First batch edited",
        },
        {
          kind: "replaceTextRunPlainText",
          handle: secondRun.handle,
          text: "Second batch edited",
        },
        {
          kind: "setTextRunProperties",
          handle: secondRun.handle,
          properties: { bold: true },
        },
      ]);
      expect(renderCalls.at(-1)).toEqual([1, 2]);
      expect(editor.slides[0]?.svg).toContain("First batch edited");
      expect(editor.slides[1]?.svg).toContain("Second batch edited");

      await editor.undo();
      expect(renderCalls.at(-1)).toEqual([1, 2]);
      expect(editor.slides[0]?.svg).toContain("First");
      expect(editor.slides[1]?.svg).toContain("Second edited");

      await editor.redo();
      expect(renderCalls.at(-1)).toEqual([1, 2]);
      expect(editor.slides[0]?.svg).toContain("First batch edited");
      expect(editor.slides[1]?.svg).toContain("Second batch edited");
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("isolates shared media replacements without warnings", async () => {
    const editor = await createPptxEditorSession(await buildImageFixture(), {
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

    expect(result.warnings).toBeUndefined();
    expect(mediaBytes(editor.document, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(editor.document, "ppt/media/image2.png")).toEqual(BLUE_PNG);
    expect(
      editor.shapes(1).find((shape) => shape.name === "Shared Picture A")?.editableImageReplacement,
    ).toMatchObject({ mediaPartPath: "ppt/media/image2.png", sharedReferenceCount: 1 });
    expect(
      editor.shapes(1).find((shape) => shape.name === "Shared Picture B")?.editableImageReplacement,
    ).toMatchObject({ mediaPartPath: "ppt/media/image1.png", sharedReferenceCount: 1 });
  });

  it("applies picture crop, rerenders the target slide, and saves the srcRect", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildImageFixture(), {
        skipSystemFonts: true,
      });
      const image = editor.shapes(1).find((shape) => shape.kind === "image");
      if (image?.handle === undefined) throw new Error("image handle not found");
      const beforeSvg = editor.slides[0]?.svg;

      await editor.apply({
        kind: "setPictureCrop",
        handle: image.handle,
        left: asOoxmlPercent(25000),
        top: asOoxmlPercent(10000),
      });

      expect(renderCalls).toEqual([undefined, [1]]);
      expect(editor.slides[0]?.svg).not.toBe(beforeSvg);
      const saved = readPptx(editor.save().pptx);
      expect(saved.slides[0]?.shapes.find((shape) => shape.kind === "image")?.crop).toEqual({
        left: 25000,
        top: 10000,
      });
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("rerenders only the picture owner after copy-on-write replacement", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(
        await buildImageFixture({ includeSecondSlide: true }),
        { skipSystemFonts: true },
      );
      const image = editor.shapes(1).find((shape) => shape.kind === "image");
      if (image?.handle === undefined) throw new Error("image handle not found");

      await editor.apply({ kind: "replaceImage", handle: image.handle, bytes: BLUE_PNG });

      expect(renderCalls).toEqual([undefined, [1]]);
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("adds, selects, edits, moves, resizes, saves, deletes, undoes, and redoes a text box", async () => {
    const editor = await createPptxEditorSession(await buildShapeFixture(), {
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
    const addedRunHandle = addedShape.textBody?.paragraphs[0]?.runs[0]?.handle;
    if (addedRunHandle === undefined) throw new Error("added text run handle not found");

    await editor.applyAll([
      {
        kind: "replaceTextRunPlainText",
        handle: addedRunHandle,
        text: "Added edited",
      },
    ]);
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
    const editor = await createPptxEditorSession(await buildShapeFixture(), {
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
    const editor = await createPptxEditorSession(
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
    const editor = await createPptxEditorSession(
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
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildTwoSlideFixture(), {
        skipSystemFonts: true,
      });
      const firstSlide = editor.slides[0];
      if (firstSlide?.handle === undefined) throw new Error("first slide handle not found");

      const duplicated = await editor.apply({ kind: "duplicateSlide", handle: firstSlide.handle });
      expect(renderCalls).toEqual([undefined, [2]]);
      expect(duplicated.slides).toHaveLength(3);
      expect(duplicated.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual(
        [
          [1, "ppt/slides/slide1.xml"],
          [2, "ppt/slides/slide3.xml"],
          [3, "ppt/slides/slide2.xml"],
        ],
      );
      expect(duplicated.slides[0]?.svg).toContain("First");
      expect(duplicated.slides[1]?.svg).toContain("First");
      expect(duplicated.history).toMatchObject({ canUndo: true, undoDepth: 1 });

      const moved = await editor.apply({
        kind: "moveSlide",
        handle: firstSlide.handle,
        toIndex: 2,
      });
      expect(renderCalls).toEqual([undefined, [2]]);
      expect(moved.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
        [1, "ppt/slides/slide3.xml"],
        [2, "ppt/slides/slide2.xml"],
        [3, "ppt/slides/slide1.xml"],
      ]);

      const duplicateSlide = moved.slides[0];
      if (duplicateSlide?.handle === undefined) throw new Error("duplicate slide handle not found");
      const deleted = await editor.apply({ kind: "deleteSlide", handle: duplicateSlide.handle });
      expect(renderCalls).toEqual([undefined, [2]]);
      expect(deleted.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
        [1, "ppt/slides/slide2.xml"],
        [2, "ppt/slides/slide1.xml"],
      ]);
      expect(deleted.history.undoDepth).toBe(3);

      expect((await editor.undo()).slides.map((slide) => slide.handle?.partPath)).toEqual([
        "ppt/slides/slide3.xml",
        "ppt/slides/slide2.xml",
        "ppt/slides/slide1.xml",
      ]);
      expect(renderCalls.at(-1)).toEqual([1]);
      const renderCallCountAfterUndo = renderCalls.length;
      expect((await editor.redo()).slides.map((slide) => slide.handle?.partPath)).toEqual([
        "ppt/slides/slide2.xml",
        "ppt/slides/slide1.xml",
      ]);
      expect(renderCalls).toHaveLength(renderCallCountAfterUndo);

      const added = await editor.apply({
        kind: "addEmptySlideFromLayout",
        layoutPartPath: "ppt/slideLayouts/slideLayout1.xml",
      });
      expect(renderCalls.at(-1)).toEqual([3]);
      expect(added.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
        [1, "ppt/slides/slide2.xml"],
        [2, "ppt/slides/slide1.xml"],
        [3, "ppt/slides/slide3.xml"],
      ]);
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("falls back to rendering all slides when a command change cannot be scoped safely", async () => {
    const before = readPptx(await buildTwoSlideFixture());
    const after = {
      ...before,
      diagnostics: [
        ...before.diagnostics,
        { severity: "warning" as const, code: "unknown-change", message: "unknown change" },
      ],
    };

    expect(affectedSlidePartPaths(before, after)).toBeUndefined();

    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionAffectedSlidesResolver(() => undefined);
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildTwoSlideFixture(), {
        skipSystemFonts: true,
      });
      const run = editor.shapes(2)[0]?.textBody?.paragraphs[0]?.runs[0];
      if (run?.handle === undefined) throw new Error("text run handle not found");

      await editor.apply({
        kind: "replaceTextRunPlainText",
        handle: run.handle,
        text: "Fallback edited",
      });

      expect(renderCalls).toEqual([undefined, undefined]);
      expect(editor.slides[0]?.svg).toContain("First");
      expect(editor.slides[1]?.svg).toContain("Fallback edited");
    } finally {
      configurePptxEditorSessionAffectedSlidesResolver(affectedSlidePartPaths);
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("scopes and renders inherited master and layout background changes", async () => {
    const created = createPptx();
    const firstSlide = created.slides[0];
    const firstLayout = created.slideLayouts[0];
    const master = created.slideMasters[0];
    if (
      firstSlide === undefined ||
      firstLayout?.handle === undefined ||
      master?.handle === undefined
    ) {
      throw new Error("createPptx should create a slide, layout, and master");
    }
    let before = addSlideLayout(created, master.handle, { name: "Second layout" });
    const secondLayout = before.slideLayouts.at(-1);
    if (secondLayout === undefined) throw new Error("second layout was not authored");
    before = addEmptySlideFromLayout(before, { layoutPartPath: secondLayout.partPath });
    const secondSlide = before.slides.at(-1);
    if (secondSlide?.handle === undefined) throw new Error("second slide was not authored");
    before = setBackground(before, secondSlide.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "FFFFFF" },
    });

    const afterMaster = setBackground(before, master.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    expect(affectedSlidePartPaths(before, afterMaster)).toEqual(new Set([firstSlide.partPath]));
    const afterMasterImage = setBackground(before, master.handle, {
      kind: "image",
      bytes: RED_PNG,
    });
    expect(affectedSlidePartPaths(before, afterMasterImage)).toEqual(
      new Set([firstSlide.partPath]),
    );
    const renderedMaster = await renderPptxSourceModelToSvg(afterMaster, {
      skipSystemFonts: true,
    });
    expect(renderedMaster.slides[0]?.svg).toContain('fill="#112233"');
    expect(renderedMaster.slides[1]?.svg).toContain('fill="#ffffff"');

    const afterLayout = setBackground(before, firstLayout.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "445566" },
    });
    expect(affectedSlidePartPaths(before, afterLayout)).toEqual(new Set([firstSlide.partPath]));
    const rendered = await renderPptxSourceModelToSvg(afterLayout, { skipSystemFonts: true });
    expect(rendered.slides[0]?.svg).toContain('fill="#445566"');
    expect(rendered.slides[1]?.svg).toContain('fill="#ffffff"');
  });

  it("resolves affected slides for layout and master shape property edits", async () => {
    const source = readPptx(await buildLayoutCatalogFixture());
    const layout = source.slideLayouts.find(
      (candidate) => candidate.partPath === "ppt/slideLayouts/slideLayout2.xml",
    );
    const master = source.slideMasters.find(
      (candidate) => candidate.partPath === "ppt/slideMasters/slideMaster1.xml",
    );
    if (layout?.handle === undefined || master?.handle === undefined) {
      throw new Error("layout or master handle not found");
    }
    const withLayoutShape = addShape(source, layout.handle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
    });
    const layoutShape = withLayoutShape.slideLayouts
      .find((candidate) => candidate.partPath === layout.partPath)
      ?.shapes.at(-1);
    if (layoutShape?.handle === undefined) throw new Error("layout shape handle not found");
    const editedLayout = setShapeFill(withLayoutShape, layoutShape.handle, { kind: "none" });

    expect([...affectedSlidePartPaths(withLayoutShape, editedLayout)!]).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
    ]);

    const withMasterShape = addShape(source, master.handle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
    });
    const masterShape = withMasterShape.slideMasters
      .find((candidate) => candidate.partPath === master.partPath)
      ?.shapes.at(-1);
    if (masterShape?.handle === undefined) throw new Error("master shape handle not found");
    const editedMaster = setShapeFill(withMasterShape, masterShape.handle, { kind: "none" });

    expect([...affectedSlidePartPaths(withMasterShape, editedMaster)!]).toEqual([
      "ppt/slides/slide1.xml",
    ]);
  });

  it("resolves all affected slides when applyAll changes inherited and slide-local content", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildTwoSlideFixture(), {
        skipSystemFonts: true,
      });
      const layoutHandle = editor.document.slideLayouts[0]?.handle;
      const secondRun = editor.shapes(2)[0]?.textBody?.paragraphs[0]?.runs[0];
      if (layoutHandle === undefined || secondRun?.handle === undefined) {
        throw new Error("layout or text run handle not found");
      }

      await editor.applyAll([
        {
          kind: "addTextBox",
          slideHandle: layoutHandle,
          offsetX: asEmu(0),
          offsetY: asEmu(0),
          width: asEmu(914400),
          height: asEmu(914400),
          text: "Inherited edit",
        },
        {
          kind: "replaceTextRunPlainText",
          handle: secondRun.handle,
          text: "Second local edit",
        },
      ]);

      expect(renderCalls).toEqual([undefined, [1, 2]]);
      expect(editor.slides[0]?.svg).toContain(">Inher</");
      expect(editor.slides[1]?.svg).toContain("Second local edit");

      await editor.undo();
      expect(renderCalls).toEqual([undefined, [1, 2], [1, 2]]);
      expect(editor.slides[0]?.svg).not.toContain(">Inher</");
      expect(editor.slides[1]?.svg).toContain("Second");
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("rejects deleting the last slide without changing the editor state", async () => {
    const editor = await createPptxEditorSession(await buildShapeFixture(), {
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

  it("preserves headless failure codes, messages, and causes in typed errors", async () => {
    const editor = await createPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const shape = editor.shapes(1)[0];
    if (shape?.handle === undefined) throw new Error("shape handle not found");
    const missingHandle = {
      ...shape.handle,
      nodeId: asSourceNodeId("missing-shape"),
    };

    const error = await capturePptxEditorError(
      editor.apply({ kind: "deleteShape", handle: missingHandle }),
    );

    expect(error).toMatchObject({
      name: "PptxEditorError",
      code: "invalid-command",
      message: "deleteShape: shape handle was not found in PptxSourceModel source",
    });
    expect(error.cause).toBeInstanceOf(Error);
    if (!(error.cause instanceof Error)) throw new Error("expected Error cause");
    expect(error.cause.message).toBe(error.message);
    expect(error.stack).toContain("PptxEditorError");
    expect(isPptxEditorError(error)).toBe(true);
    expect(editor.document.edits).toBeUndefined();
    expect(editor.history).toMatchObject({ canUndo: false, canRedo: false });
  });

  it("wraps read, render, and write integration failures with their causes", async () => {
    const readError = await capturePptxEditorError(
      createPptxEditorSession(new Uint8Array([0x00, 0x01, 0x02])),
    );
    expectErrorCodeAndCause(readError, "read-failed");

    const renderCause = new Error("renderer unavailable");
    configurePptxEditorSessionRenderer(() => Promise.reject(renderCause));
    try {
      const initialRenderError = await capturePptxEditorError(
        createPptxEditorSession(await buildShapeFixture(), { skipSystemFonts: true }),
      );
      expectErrorCodeAndCause(initialRenderError, "render-failed", renderCause);
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }

    let renderCount = 0;
    configurePptxEditorSessionRenderer(async (source, options) => {
      renderCount += 1;
      if (renderCount > 1) throw renderCause;
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildShapeFixture(), {
        skipSystemFonts: true,
      });
      const shape = editor.shapes(1)[0];
      if (shape?.handle === undefined) throw new Error("shape handle not found");
      const renderError = await capturePptxEditorError(
        editor.apply({ kind: "deleteShape", handle: shape.handle }),
      );
      expectErrorCodeAndCause(renderError, "render-failed", renderCause);
      expect(editor.shapes(1)).toHaveLength(0);
      expect(editor.history).toMatchObject({ canUndo: true, undoDepth: 1 });
      expect(editor.slides[0]?.svg).toContain("Original");
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }

    const editor = await createPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const run = editor.shapes(1)[0]?.textBody?.paragraphs[0]?.runs[0];
    if (run?.handle === undefined) throw new Error("text run handle not found");
    await editor.apply({ kind: "replaceTextRunPlainText", handle: run.handle, text: "Edited" });
    const edit = editor.document.edits?.[0];
    if (edit === undefined) throw new Error("expected an edit");
    Object.defineProperty(editor.document, "edits", { value: [edit, edit] });

    const writeError = captureSynchronousPptxEditorError(() => editor.save());
    expectErrorCodeAndCause(writeError, "write-failed");
  });

  it("reports the same operation failure code from Node and browser entries", async () => {
    const nodeEditor = await createPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const browserEntry = await import("./browser.js");
    const browserEditor = await browserEntry.createPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });

    for (const editor of [nodeEditor, browserEditor]) {
      const shape = editor.shapes(1)[0];
      if (shape?.handle === undefined) throw new Error("shape handle not found");
      const error = captureSynchronousPptxEditorError(() =>
        editor.selectShape({ ...shape.handle, nodeId: asSourceNodeId("missing-shape") }),
      );
      expect(error.code).toBe("invalid-selection");
      expect(browserEntry.isPptxEditorError(error)).toBe(true);
    }
  });

  it("exposes the same ordered master and layout catalog from Node and browser entries", async () => {
    const input = await buildLayoutCatalogFixture();
    const browserEntry = await import("./browser.js");
    const editors = [
      await createPptxEditorSession(input, { skipSystemFonts: true }),
      await browserEntry.createPptxEditorSession(input, { skipSystemFonts: true }),
    ];

    const expected = [
      {
        handle: { partPath: "ppt/slideMasters/slideMaster2.xml" },
        name: "Second Master",
        layouts: [
          {
            handle: { partPath: "ppt/slideLayouts/slideLayout3.xml" },
            name: "Hidden Layout",
            type: "title",
            hidden: true,
            slideReferenceCount: 0,
          },
          {
            handle: { partPath: "ppt/slideLayouts/slideLayout2.xml" },
            name: "Popular Layout",
            type: "twoObj",
            hidden: false,
            slideReferenceCount: 2,
          },
        ],
      },
      {
        handle: { partPath: "ppt/slideMasters/slideMaster1.xml" },
        name: "First Master",
        layouts: [
          {
            handle: { partPath: "ppt/slideLayouts/slideLayout1.xml" },
            name: "Visible by Default",
            type: "blank",
            hidden: false,
            slideReferenceCount: 1,
          },
        ],
      },
    ];

    for (const editor of editors) {
      expect(editor.layoutCatalog).toEqual(expected);
    }
  });

  it("groups and ungroups with topology and selection restored by history", async () => {
    const editor = await createPptxEditorSession(buildGroupCommandFixture(), {
      skipSystemFonts: true,
    });
    const [first, second] = editor.shapes(1).map((shape) => shape.handle);
    if (first === undefined || second === undefined) {
      throw new Error("group command fixture handles are missing");
    }
    editor.selectShape(second);

    const grouped = await editor.groupShapes([first, second]);
    const group = editor.shapes(1)[0];
    if (group?.kind !== "group" || group.handle === undefined) {
      throw new Error("high-level group command did not create a group");
    }
    expect(grouped.selection).toEqual({ shapeHandle: group.handle });
    expect(grouped.history.undoDepth).toBe(1);
    expect(grouped.slides[0]?.svg).toContain('aria-label="Group 4"');

    expect((await editor.undo()).selection).toEqual({ shapeHandle: second });
    expect(editor.shapes(1).map((shape) => shape.kind)).toEqual(["shape", "shape", "shape"]);
    expect((await editor.redo()).selection).toEqual({ shapeHandle: group.handle });

    const ungrouped = await editor.ungroupShape(group.handle);
    expect(ungrouped.selection).toEqual({ shapeHandle: first });
    expect(editor.shapes(1).map((shape) => shape.kind)).toEqual(["shape", "shape", "shape"]);
    expect((await editor.undo()).selection).toEqual({ shapeHandle: group.handle });
    expect((await editor.redo()).selection).toEqual({ shapeHandle: first });
  });

  it("exposes the same integrated group command behavior from Node and browser entries", async () => {
    const browserEntry = await import("./browser.js");
    const input = buildGroupCommandFixture();
    const editors = [
      await createPptxEditorSession(input, { skipSystemFonts: true }),
      await browserEntry.createPptxEditorSession(input, { skipSystemFonts: true }),
    ];

    for (const editor of editors) {
      const handles = editor
        .shapes(1)
        .slice(0, 2)
        .map((shape) => shape.handle);
      if (handles[0] === undefined || handles[1] === undefined) {
        throw new Error("group command fixture handles are missing");
      }
      const result = await editor.apply({ kind: "groupShapes", shapeHandles: handles });
      const group = editor.shapes(1)[0];
      if (group?.kind !== "group" || group.handle === undefined) {
        throw new Error("integrated group command did not create a selectable group");
      }
      expect(result.selection).toEqual({ shapeHandle: group.handle });
      expect(editor.document.slides[0]?.shapes[0]?.kind).toBe("group");
    }
  });

  it("recognizes a structurally valid editor error from another realm", () => {
    const foreignError = {
      name: "PptxEditorError",
      code: "render-failed",
      message: "foreign renderer failed",
      cause: new Error("foreign cause"),
    };

    expect(isPptxEditorError(foreignError)).toBe(true);
    expect(isPptxEditorError({ ...foreignError, code: "unknown" })).toBe(false);
    expect(isPptxEditorError({ ...foreignError, name: "Error" })).toBe(false);
  });

  it("counts unparsed image relationships in image replacement metadata", async () => {
    const editor = await createPptxEditorSession(
      await buildImageFixture({ includeUnusedImageRelationship: true }),
      { skipSystemFonts: true },
    );
    const image = editor.shapes(1).find((shape) => shape.kind === "image");

    expect(image?.editableImageReplacement?.sharedReferenceCount).toBe(3);
  });

  it("counts preserved image fills in image replacement metadata", async () => {
    const editor = await createPptxEditorSession(
      await buildImageFixture({ includeImageFill: true }),
      { skipSystemFonts: true },
    );
    const image = editor.shapes(1).find((shape) => shape.kind === "image");

    expect(image?.editableImageReplacement?.sharedReferenceCount).toBe(3);
  });
});

function buildGroupCommandFixture(): Uint8Array {
  let source = createPptx();
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("group command slide handle is missing");
  for (const offsetX of [914400, 2743200, 4572000]) {
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(offsetX),
      offsetY: asEmu(914400),
      width: asEmu(1371600),
      height: asEmu(914400),
    });
  }
  return writePptx(source);
}

async function buildLayoutCatalogFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        [1, 2, 3]
          .map(
            (number) =>
              `<Override PartName="/ppt/slides/slide${String(number)}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
          )
          .join("") +
        [1, 2, 3]
          .map(
            (number) =>
              `<Override PartName="/ppt/slideLayouts/slideLayout${String(number)}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`,
          )
          .join("") +
        [1, 2]
          .map(
            (number) =>
              `<Override PartName="/ppt/slideMasters/slideMaster${String(number)}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`,
          )
          .join("") +
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
        `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster2"/><p:sldMasterId id="2147483649" r:id="rIdMaster1"/></p:sldMasterIdLst>` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/><p:sldId id="258" r:id="rIdSlide3"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<Relationship Id="rIdMaster2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>` +
        [1, 2, 3]
          .map(
            (number) =>
              `<Relationship Id="rIdSlide${String(number)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${String(number)}.xml"/>`,
          )
          .join("") +
        `</Relationships>`,
    ),
  );

  zip.file("ppt/slideMasters/slideMaster1.xml", slideMasterXml("First Master", [[1, 1]]));
  zip.file(
    "ppt/slideMasters/slideMaster2.xml",
    slideMasterXml("Second Master", [
      [3, 3],
      [2, 2],
    ]),
  );
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", masterLayoutRels([1]));
  zip.file("ppt/slideMasters/_rels/slideMaster2.xml.rels", masterLayoutRels([2, 3]));

  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml("Visible by Default", "blank"));
  zip.file("ppt/slideLayouts/slideLayout2.xml", slideLayoutXml("Popular Layout", "twoObj", true));
  zip.file("ppt/slideLayouts/slideLayout3.xml", slideLayoutXml("Hidden Layout", "title", false));
  for (const [layoutNumber, masterNumber] of [
    [1, 1],
    [2, 2],
    [3, 2],
  ] as const) {
    zip.file(
      `ppt/slideLayouts/_rels/slideLayout${String(layoutNumber)}.xml.rels`,
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster${String(masterNumber)}.xml"/>` +
          `</Relationships>`,
      ),
    );
  }

  for (const [slideNumber, layoutNumber] of [
    [1, 1],
    [2, 2],
    [3, 2],
  ] as const) {
    zip.file(
      `ppt/slides/slide${String(slideNumber)}.xml`,
      xml(
        `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>`,
      ),
    );
    zip.file(
      `ppt/slides/_rels/slide${String(slideNumber)}.xml.rels`,
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${String(layoutNumber)}.xml"/>` +
          `</Relationships>`,
      ),
    );
  }

  return zip.generateAsync({ type: "uint8array" });
}

function slideMasterXml(name: string, layouts: readonly (readonly [number, number])[]): Uint8Array {
  return xml(
    `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<p:cSld name="${name}"><p:spTree/></p:cSld>` +
      `<p:sldLayoutIdLst>${layouts
        .map(
          ([id, relationshipNumber]) =>
            `<p:sldLayoutId id="${String(2147483648 + id)}" r:id="rIdLayout${String(relationshipNumber)}"/>`,
        )
        .join("")}</p:sldLayoutIdLst>` +
      `</p:sldMaster>`,
  );
}

function masterLayoutRels(layoutNumbers: readonly number[]): Uint8Array {
  return xml(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${layoutNumbers
      .map(
        (number) =>
          `<Relationship Id="rIdLayout${String(number)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${String(number)}.xml"/>`,
      )
      .join("")}</Relationships>`,
  );
}

function slideLayoutXml(name: string, type: string, show?: boolean): Uint8Array {
  return xml(
    `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="${type}"${show === undefined ? "" : ` show="${show ? "1" : "0"}"`}>` +
      `<p:cSld name="${name}"><p:spTree/></p:cSld>` +
      `</p:sldLayout>`,
  );
}

async function capturePptxEditorError(promise: Promise<unknown>): Promise<PptxEditorError> {
  try {
    await promise;
  } catch (error) {
    if (isPptxEditorError(error)) return error;
    throw error;
  }
  throw new Error("expected PptxEditorError");
}

function captureSynchronousPptxEditorError(operation: () => unknown): PptxEditorError {
  try {
    operation();
  } catch (error) {
    if (isPptxEditorError(error)) return error;
    throw error;
  }
  throw new Error("expected PptxEditorError");
}

function expectErrorCodeAndCause(
  error: PptxEditorError,
  code: PptxEditorErrorCode,
  cause?: unknown,
): void {
  expect(error.code).toBe(code);
  expect(error.name).toBe("PptxEditorError");
  if (cause === undefined) {
    expect(error.cause).toBeDefined();
  } else {
    expect(error.cause).toBe(cause);
  }
}

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
  options: {
    readonly includeImageFill?: boolean;
    readonly includeUnusedImageRelationship?: boolean;
    readonly includeSecondSlide?: boolean;
  } = {},
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
        (options.includeSecondSlide === true
          ? `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
          : "") +
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
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/>` +
        (options.includeSecondSlide === true ? `<p:sldId id="257" r:id="rIdSlide2"/>` : "") +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        (options.includeSecondSlide === true
          ? `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>`
          : "") +
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
        (options.includeImageFill === true
          ? `<p:sp><p:nvSpPr><p:cNvPr id="22" name="Shared Fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
            `<p:spPr><a:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
            `<a:prstGeom prst="rect"/></p:spPr></p:sp>`
          : "") +
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
  if (options.includeSecondSlide === true) {
    zip.file("ppt/slides/slide2.xml", zip.file("ppt/slides/slide1.xml")?.async("uint8array"));
    zip.file(
      "ppt/slides/_rels/slide2.xml.rels",
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
          `</Relationships>`,
      ),
    );
  }
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
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
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
  const slideLayoutRelationship = xml(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `</Relationships>`,
  );
  zip.file("ppt/slides/_rels/slide1.xml.rels", slideLayoutRelationship);
  zip.file("ppt/slides/_rels/slide2.xml.rels", slideLayoutRelationship);
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
  );

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
