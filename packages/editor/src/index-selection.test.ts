import {
  asEmu,
  asPartPath,
  asSourceNodeId,
  findShapeNodeBySourceHandle,
  readPptx,
  type SourceHandle,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";
import {
  buildTextEditFixture,
  createThreeShapeSource,
  expectApplied,
  expectHistory,
  firstShape,
  requireHandle,
} from "./index.test-helpers.js";

describe("EditorSession selection", () => {
  it("selects and deselects a shape without changing undo history", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);

    const selected = session.selectShape(handle);

    expect(selected).toEqual({
      ok: true,
      selection: { shapeHandle: handle },
    });
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);

    session.deselectShape();

    expect(session.selection).toBeUndefined();
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("rejects missing shape selection without changing selection", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);
    const missingHandle = {
      partPath: asPartPath("ppt/slides/slide1.xml"),
      nodeId: asSourceNodeId("999"),
      orderingSlot: 99,
    } satisfies SourceHandle;

    expect(session.selectShape(handle)).toMatchObject({ ok: true });
    const rejected = session.selectShape(missingHandle);

    expect(rejected).toMatchObject({
      ok: false,
      code: "invalid-selection",
    });
    expect(rejected.ok).toBe(false);
    if (!rejected.ok) {
      expect(rejected.message).toMatch(/shape handle was not found/);
    }
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("keeps shape selection across move and resize edits, undo, and redo", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);

    expect(session.selectShape(handle)).toMatchObject({ ok: true });
    expectApplied(
      session.apply({
        kind: "moveShape",
        handle,
        offsetX: asEmu(1000),
        offsetY: asEmu(2000),
      }),
    );
    expectApplied(
      session.apply({
        kind: "resizeShape",
        handle,
        width: asEmu(3000),
        height: asEmu(4000),
      }),
    );

    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();

    expectHistory(session.undo());
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();

    expectHistory(session.undo());
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();

    expectHistory(session.redo());
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();

    expectHistory(session.redo());
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();
  });

  it("clears selection when the selected shape is deleted", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);

    expect(session.selectShape(handle)).toMatchObject({ ok: true });
    expectApplied(session.apply({ kind: "deleteShape", handle }));

    expect(session.selection).toBeUndefined();
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeUndefined();
    expectHistory(session.undo());
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();
    expect(session.selection).toBeUndefined();
    expectHistory(session.redo());
    expect(findShapeNodeBySourceHandle(session.document, handle)).toBeUndefined();
    expect(session.selection).toBeUndefined();
  });

  it("selects group and first ungrouped child while undo and redo restore selection", () => {
    const source = createThreeShapeSource();
    const handles = source.slides[0]?.shapes.map((shape) => requireHandle(shape.handle)) ?? [];
    const session = createEditorSession(source);
    expect(session.selectShape(handles[1])).toMatchObject({ ok: true });

    expectApplied(session.apply({ kind: "groupShapes", shapeHandles: handles.slice(0, 2) }));
    const group = session.document.slides[0]?.shapes[0];
    expect(group?.kind).toBe("group");
    if (group?.kind !== "group" || group.handle === undefined) {
      throw new Error("group command did not create a selectable group");
    }
    expect(group.children.map((child) => child.nodeId)).toEqual(
      handles.slice(0, 2).map((handle) => handle.nodeId),
    );
    expect(session.selection).toEqual({ shapeHandle: group.handle });

    expectHistory(session.undo());
    expect(session.document).toBe(source);
    expect(session.selection).toEqual({ shapeHandle: handles[1] });
    expectHistory(session.redo());
    expect(session.selection).toEqual({ shapeHandle: group.handle });

    expectApplied(session.apply({ kind: "ungroupShape", groupHandle: group.handle }));
    expect(session.document.slides[0]?.shapes.map((shape) => shape.nodeId)).toEqual(
      handles.map((handle) => handle.nodeId),
    );
    expect(session.selection).toEqual({ shapeHandle: handles[0] });
    expectHistory(session.undo());
    expect(session.selection).toEqual({ shapeHandle: group.handle });
    expectHistory(session.redo());
    expect(session.selection).toEqual({ shapeHandle: handles[0] });
  });

  it("rejects invalid grouping without changing document, selection, or history", () => {
    const source = createThreeShapeSource();
    const handles = source.slides[0]?.shapes.map((shape) => requireHandle(shape.handle)) ?? [];
    const session = createEditorSession(source);
    expect(session.selectShape(handles[1])).toMatchObject({ ok: true });
    expectApplied(
      session.apply({
        kind: "moveShape",
        handle: handles[1],
        offsetX: asEmu(2500),
        offsetY: asEmu(0),
      }),
    );
    expectHistory(session.undo());
    const before = session.document;
    const selection = session.selection;

    const rejected = session.apply({
      kind: "groupShapes",
      shapeHandles: [handles[0], handles[2]],
    });
    const rejectedUngroup = session.apply({
      kind: "ungroupShape",
      groupHandle: handles[0],
    });

    expect(rejected).toMatchObject({ ok: false, code: "invalid-command" });
    expect(rejectedUngroup).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(session.selection).toBe(selection);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(1);
  });

  it("moves shapes across identity-mapped parents with selection and history preserved", () => {
    const source = createThreeShapeSource();
    const handles = source.slides[0]?.shapes.map((shape) => requireHandle(shape.handle)) ?? [];
    const session = createEditorSession(source);
    expectApplied(session.apply({ kind: "groupShapes", shapeHandles: handles.slice(0, 2) }));
    const group = session.document.slides[0]?.shapes[0];
    if (group?.kind !== "group" || group.handle === undefined) {
      throw new Error("move command fixture group is missing");
    }
    expect(session.selectShape(handles[2])).toMatchObject({ ok: true });

    expectApplied(
      session.apply({
        kind: "moveShapes",
        shapeHandles: [handles[2]],
        destinationHandle: group.handle,
        beforeShapeHandle: handles[0],
      }),
    );
    expect(session.selection).toEqual({ shapeHandle: handles[2] });
    expect(session.document.slides[0]?.shapes).toHaveLength(1);
    expect(
      session.document.slides[0]?.shapes[0]?.kind === "group"
        ? session.document.slides[0].shapes[0].children.map((shape) => shape.nodeId)
        : [],
    ).toEqual([handles[2].nodeId, handles[0].nodeId, handles[1].nodeId]);

    expectHistory(session.undo());
    expect(session.document.slides[0]?.shapes).toHaveLength(2);
    expect(session.selection).toEqual({ shapeHandle: handles[2] });
    expectHistory(session.redo());
    expect(session.document.slides[0]?.shapes).toHaveLength(1);
    expect(session.selection).toEqual({ shapeHandle: handles[2] });

    const before = session.document;
    const undoDepth = session.undoDepth;
    const rejected = session.applyAll([
      {
        kind: "moveShapes",
        shapeHandles: [handles[2]],
        destinationHandle: requireHandle(session.document.slides[0]?.handle),
      },
      {
        kind: "moveShapes",
        shapeHandles: [handles[0], handles[2]],
        destinationHandle: requireHandle(session.document.slides[0]?.handle),
      },
    ]);
    expect(rejected).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(undoDepth);
    expect(session.selection).toEqual({ shapeHandle: handles[2] });
  });

  it("keeps a later selection across ordinary command undo and redo", () => {
    const source = createThreeShapeSource();
    const handles = source.slides[0]?.shapes.map((shape) => requireHandle(shape.handle)) ?? [];
    const session = createEditorSession(source);
    expect(session.selectShape(handles[0])).toMatchObject({ ok: true });
    expectApplied(
      session.apply({
        kind: "moveShape",
        handle: handles[0],
        offsetX: asEmu(500),
        offsetY: asEmu(500),
      }),
    );
    expect(session.selectShape(handles[2])).toMatchObject({ ok: true });

    expectHistory(session.undo());
    expect(session.selection).toEqual({ shapeHandle: handles[2] });
    expectHistory(session.redo());
    expect(session.selection).toEqual({ shapeHandle: handles[2] });
  });

  it("rejects unsupported group child kinds and empty ungroup targets atomically", () => {
    const unsupportedSource = createThreeShapeSource();
    const unsupportedHandles =
      unsupportedSource.slides[0]?.shapes.map((shape) => requireHandle(shape.handle)) ?? [];
    Object.defineProperty(unsupportedSource.slides[0]?.shapes[0], "kind", {
      value: "smartArt",
    });
    const unsupportedSession = createEditorSession(unsupportedSource);
    const unsupportedBefore = unsupportedSession.document;
    expect(
      unsupportedSession.apply({
        kind: "groupShapes",
        shapeHandles: unsupportedHandles.slice(0, 2),
      }),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(unsupportedSession.document).toBe(unsupportedBefore);
    expect(unsupportedSession.undoDepth).toBe(0);

    const source = createThreeShapeSource();
    const handles = source.slides[0]?.shapes.map((shape) => requireHandle(shape.handle)) ?? [];
    const session = createEditorSession(source);
    expectApplied(session.apply({ kind: "groupShapes", shapeHandles: handles.slice(0, 2) }));
    const group = session.document.slides[0]?.shapes[0];
    if (group?.kind !== "group" || group.handle === undefined) {
      throw new Error("empty group rejection fixture is missing");
    }
    Object.defineProperty(group, "children", { value: [] });
    const before = session.document;
    const selection = session.selection;
    const undoDepth = session.undoDepth;

    expect(session.apply({ kind: "ungroupShape", groupHandle: group.handle })).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(session.document).toBe(before);
    expect(session.selection).toBe(selection);
    expect(session.undoDepth).toBe(undoDepth);
    expect(session.redoDepth).toBe(0);
  });

  it("classifies ambiguous group command handles as atomic invalid-command failures", () => {
    const source = createThreeShapeSource();
    const first = source.slides[0]?.shapes[0];
    const second = source.slides[0]?.shapes[1];
    const third = source.slides[0]?.shapes[2];
    if (
      first?.handle?.nodeId === undefined ||
      second?.handle === undefined ||
      third?.handle === undefined
    ) {
      throw new Error("ambiguous group command fixture handles are missing");
    }
    Object.defineProperty(second.handle, "nodeId", { value: first.handle.nodeId });
    const session = createEditorSession(source);
    const before = session.document;

    const result = session.apply({
      kind: "groupShapes",
      shapeHandles: [first.handle, third.handle],
    });

    expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.message).toContain("duplicate node id");
    expect(session.document).toBe(before);
    expect(session.selection).toBeUndefined();
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });
});
