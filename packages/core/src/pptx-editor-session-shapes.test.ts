import {
  asEmu,
  createPptx,
  createPptxAuthoringSession,
  findShapeNodeBySourceHandle,
  groupShapes,
  readPptx,
  writePptx,
} from "@pptx-glimpse/document";
import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { createPptxEditorSession } from "./index.js";
import {
  buildDrawingDeleteSessionFixture,
  buildExternallyConnectedGroupFixture,
  buildGroupCommandFixture,
  buildNestedDrawingDeleteSessionFixture,
  buildShapeFixture,
  connectorByName,
  handleKey,
  shapeByText,
} from "./pptx-editor-session.test-helpers.js";

describe("PptxEditorSession - shapes", () => {
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

  it("does not expose group delete when an external connector references a child", async () => {
    const editor = await createPptxEditorSession(await buildExternallyConnectedGroupFixture(), {
      skipSystemFonts: true,
    });
    const editableGroup = editor.shapes(1).find((shape) => shape.kind === "group");
    if (editableGroup?.handle === undefined) throw new Error("editable group was not found");

    expect(editableGroup.editableDelete).toBeUndefined();
    await expect(editor.deleteShape(editableGroup.handle)).rejects.toMatchObject({
      code: "invalid-command",
    });
  });

  it("rerenders, saves, and restores history for picture, table, chart, and group deletes", async () => {
    for (const kind of ["image", "table", "chart", "group"] as const) {
      const editor = await createPptxEditorSession(buildDrawingDeleteSessionFixture(), {
        skipSystemFonts: true,
      });
      const target = editor.shapes(1).find((shape) => shape.kind === kind);
      if (target?.handle === undefined) throw new Error(`${kind} delete target was not found`);
      expect(target.editableDelete).toBe(true);
      editor.selectShape(target.handle);

      const deleted = await editor.deleteSelectedShape();
      expect(deleted.selection).toBeUndefined();
      expect(deleted.history.undoDepth).toBe(1);
      expect(
        editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(target.handle)),
      ).toBe(false);
      expect(
        findShapeNodeBySourceHandle(readPptx(editor.save().pptx), target.handle),
      ).toBeUndefined();

      await editor.undo();
      expect(
        editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(target.handle)),
      ).toBe(true);
      await editor.redo();
      expect(
        editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(target.handle)),
      ).toBe(false);
    }
  });

  it("rerenders, saves, selects, and restores history for nested drawing deletes", async () => {
    for (const kind of ["shape", "connector", "image", "table", "chart", "group"] as const) {
      const editor = await createPptxEditorSession(buildNestedDrawingDeleteSessionFixture(), {
        skipSystemFonts: true,
      });
      const candidates = editor.shapes(1).filter((shape) => shape.kind === kind);
      const target = candidates.at(-1);
      if (target?.handle === undefined) throw new Error(`${kind} nested target was not found`);
      expect(target.editableDelete).toBe(true);
      const beforeSvg = (await editor.renderCurrentSlides())[0]?.svg;
      editor.selectShape(target.handle);

      const deleted = await editor.deleteSelectedShape();
      expect(deleted.selection).toBeUndefined();
      expect(deleted.history.undoDepth).toBe(1);
      expect(deleted.slides[0]?.svg).not.toBe(beforeSvg);
      expect(
        editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(target.handle)),
      ).toBe(false);
      expect(
        findShapeNodeBySourceHandle(readPptx(editor.save().pptx), target.handle),
      ).toBeUndefined();

      await editor.undo();
      expect(
        editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(target.handle)),
      ).toBe(true);
      await editor.redo();
      expect(
        editor.shapes(1).some((shape) => handleKey(shape.handle) === handleKey(target.handle)),
      ).toBe(false);
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

  it("rerenders, saves, and restores an identity-mapped cross-parent move", async () => {
    const editor = await createPptxEditorSession(buildGroupCommandFixture(), {
      skipSystemFonts: true,
    });
    const [first, second, third] = editor.shapes(1).map((shape) => shape.handle);
    if (first === undefined || second === undefined || third === undefined) {
      throw new Error("cross-parent move fixture handles are missing");
    }
    await editor.groupShapes([first, second]);
    const group = editor.document.slides[0]?.shapes[0];
    if (group?.kind !== "group" || group.handle === undefined) {
      throw new Error("cross-parent move fixture group is missing");
    }
    editor.selectShape(third);
    const beforeSvg = (await editor.renderCurrentSlides())[0]?.svg;

    const moved = await editor.moveShapes([third], group.handle, { beforeShapeHandle: first });
    expect(moved.selection).toEqual({ shapeHandle: third });
    expect(moved.history.undoDepth).toBe(2);
    expect(moved.slides[0]?.svg).not.toBe(beforeSvg);
    const movedGroup = editor.document.slides[0]?.shapes[0];
    expect(
      movedGroup?.kind === "group" ? movedGroup.children.map((shape) => shape.nodeId) : [],
    ).toEqual([third.nodeId, first.nodeId, second.nodeId]);
    const savedGroup = readPptx(editor.save().pptx).slides[0]?.shapes[0];
    expect(
      savedGroup?.kind === "group" ? savedGroup.children.map((shape) => shape.nodeId) : [],
    ).toEqual([third.nodeId, first.nodeId, second.nodeId]);

    await editor.undo();
    expect(editor.document.slides[0]?.shapes.map((shape) => shape.nodeId)).toEqual([
      group.nodeId,
      third.nodeId,
    ]);
    await editor.redo();
    const redoneGroup = editor.document.slides[0]?.shapes[0];
    expect(
      redoneGroup?.kind === "group" ? redoneGroup.children.map((shape) => shape.nodeId) : [],
    ).toEqual([third.nodeId, first.nodeId, second.nodeId]);
  });

  it("rerenders, saves, and restores a representable affine cross-parent move", async () => {
    const editor = await createPptxEditorSession(buildAffineCrossParentMoveFixture(), {
      skipSystemFonts: true,
    });
    const group = editor.document.slides[0]?.shapes.find((shape) => shape.kind === "group");
    const root = editor.document.slides[0]?.shapes.find((shape) => shape.kind === "shape");
    if (group?.kind !== "group" || group.handle === undefined || root?.handle === undefined) {
      throw new Error("affine cross-parent move fixture is missing handles");
    }
    editor.selectShape(root.handle);
    const beforeTransform = root.transform;
    const beforeSvg = (await editor.renderCurrentSlides())[0]?.svg;

    const moved = await editor.moveShapes([root.handle], group.handle);
    expect(moved.selection).toEqual({ shapeHandle: root.handle });
    expect(moved.history.undoDepth).toBe(1);
    expect(moved.slides[0]?.svg).toBeDefined();
    expect(moved.slides[0]?.svg).not.toBe(beforeSvg);
    const movedGroup = editor.document.slides[0]?.shapes[0];
    const movedRoot = movedGroup?.kind === "group" ? movedGroup.children.at(-1) : undefined;
    expect(movedRoot?.transform).not.toEqual(beforeTransform);

    const saved = readPptx(editor.save().pptx);
    const savedGroup = saved.slides[0]?.shapes[0];
    expect(savedGroup?.kind === "group" ? savedGroup.children.at(-1)?.nodeId : undefined).toBe(
      root.nodeId,
    );
    await editor.undo();
    expect(editor.document.slides[0]?.shapes.at(-1)?.nodeId).toBe(root.nodeId);
    await editor.redo();
    expect(editor.document.slides[0]?.shapes[0]?.kind).toBe("group");
  });

  it("rerenders both affected slides and saves cross-slide drawing identity remaps", async () => {
    const editor = await createPptxEditorSession(buildCrossSlideMoveFixture(), {
      skipSystemFonts: true,
    });
    const sourceHandle = editor.shapes(1)[0]?.handle;
    const destinationHandle = editor.document.slides[1]?.handle;
    if (sourceHandle === undefined || destinationHandle === undefined) {
      throw new Error("cross-slide move fixture handles are missing");
    }
    editor.selectShape(sourceHandle);
    const before = await editor.renderCurrentSlides();

    const response = await editor.moveShapesAcrossSlides([sourceHandle], destinationHandle);
    const movedHandle = response.selection?.shapeHandle;
    if (movedHandle === undefined) throw new Error("cross-slide selection mapping is missing");
    expect(movedHandle.partPath).toBe(editor.document.slides[1]?.partPath);
    expect(movedHandle.nodeId).not.toBe(sourceHandle.nodeId);
    expect(response.slides[0]?.svg).not.toBe(before[0]?.svg);
    expect(response.slides[1]?.svg).not.toBe(before[1]?.svg);

    const persisted = readPptx(editor.save().pptx);
    expect(persisted.slides[0]?.shapes.map(sourceShapeName)).not.toContain("Moved");
    expect(persisted.slides[1]?.shapes.map(sourceShapeName)).toEqual(["Destination", "Moved"]);
    await editor.undo();
    expect(editor.selection).toEqual({ shapeHandle: sourceHandle });
    await editor.redo();
    expect(editor.selection).toEqual({ shapeHandle: movedHandle });
  });

  it("rerenders both affected slides and preserves chart selection through history", async () => {
    const editor = await createPptxEditorSession(buildCrossSlideChartMoveFixture(), {
      skipSystemFonts: true,
    });
    const chart = editor.document.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    const destinationHandle = editor.document.slides[1]?.handle;
    if (chart?.kind !== "chart" || chart.handle === undefined || destinationHandle === undefined) {
      throw new Error("cross-slide chart move fixture handles are missing");
    }
    editor.selectShape(chart.handle);
    const before = await editor.renderCurrentSlides();

    const response = await editor.moveShapesAcrossSlides([chart.handle], destinationHandle);
    const movedHandle = response.selection?.shapeHandle;
    if (movedHandle === undefined) throw new Error("chart selection mapping is missing");
    expect(response.slides[0]?.svg).not.toBe(before[0]?.svg);
    expect(response.slides[1]?.svg).not.toBe(before[1]?.svg);
    expect(movedHandle.partPath).toBe(editor.document.slides[1]?.partPath);
    const persisted = readPptx(editor.save().pptx);
    expect(persisted.slides[0]?.shapes.some((shape) => shape.kind === "chart")).toBe(false);
    expect(persisted.slides[1]?.shapes.some((shape) => shape.kind === "chart")).toBe(true);
    await editor.undo();
    expect(editor.selection).toEqual({ shapeHandle: chart.handle });
    await editor.redo();
    expect(editor.selection).toEqual({ shapeHandle: movedHandle });
  });
});

function buildCrossSlideMoveFixture(): Uint8Array {
  const initial = createPptx();
  const session = createPptxAuthoringSession(initial);
  const first = initial.slides[0]?.handle;
  const layout = initial.slideLayouts[0]?.handle;
  if (first === undefined || layout === undefined) {
    throw new Error("cross-slide move fixture roots are missing");
  }
  const second = session.addEmptySlideFromLayout({ layoutPartPath: layout.partPath });
  session.target(first).addShape({
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(0),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Moved",
  });
  session.target(second).addShape({
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(1200),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Destination",
  });
  return writePptx(session.source);
}

function buildCrossSlideChartMoveFixture(): Uint8Array {
  const initial = createPptx();
  const session = createPptxAuthoringSession(initial);
  const first = initial.slides[0]?.handle;
  const layout = initial.slideLayouts[0]?.handle;
  if (first === undefined || layout === undefined) {
    throw new Error("cross-slide chart move fixture roots are missing");
  }
  session.addEmptySlideFromLayout({ layoutPartPath: layout.partPath });
  session.target(first).addChart({
    chartType: "bar",
    offsetX: asEmu(0),
    offsetY: asEmu(0),
    width: asEmu(2400),
    height: asEmu(1800),
    name: "Moved chart",
    series: [{ name: "Revenue", categories: ["A", "B"], values: [10, 20] }],
  });
  return writePptx(session.source);
}

function sourceShapeName(shape: { readonly kind: string }): string | undefined {
  return "name" in shape && typeof shape.name === "string" ? shape.name : undefined;
}

function buildAffineCrossParentMoveFixture(): Uint8Array {
  const source = readPptx(buildGroupCommandFixture());
  const handles = source.slides[0]?.shapes.map((shape) => shape.handle) ?? [];
  if (handles[0] === undefined || handles[1] === undefined) {
    throw new Error("affine fixture source handles are missing");
  }
  const grouped = groupShapes(source, [handles[0], handles[1]]);
  const archive = unzipSync(writePptx(grouped));
  const partPath = "ppt/slides/slide1.xml";
  const xml = new TextDecoder().decode(requireValue(archive[partPath]));
  const groupBlock = requireValue(xml.match(/<p:grpSp\b[^>]*>[\s\S]*?<\/p:grpSp>/)?.[0]);
  const transformedGroup = groupBlock.replace(
    /<a:xfrm[\s\S]*?<\/a:xfrm>/,
    `<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/><a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm>`,
  );
  archive[partPath] = new TextEncoder().encode(xml.replace(groupBlock, transformedGroup));
  return zipSync(archive);
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("test fixture value is missing");
  return value;
}
