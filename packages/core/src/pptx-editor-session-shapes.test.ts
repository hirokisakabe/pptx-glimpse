import { asEmu, readPptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createPptxEditorSession } from "./index.js";
import {
  buildGroupCommandFixture,
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
});
