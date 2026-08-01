import {
  asEmu,
  asPartPath,
  asSourceNodeId,
  findShapeNodeBySourceHandle,
  readPptx,
  type SourceHandle,
  writePptx,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";
import {
  buildShapeStyleFixture,
  buildTextEditFixture,
  expectApplied,
  expectHistory,
  findConnectorByName,
  findShapeByName,
  firstRun,
  firstShape,
  requireHandle,
  requireShape,
  shapeWithoutTransform,
} from "./index.test-helpers.js";

describe("EditorSession shape add/delete commands", () => {
  it("adds a text box and lets existing text/xfrm commands edit it before save", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const withTextBox = expectApplied(
      session.apply({
        kind: "addTextBox",
        slideHandle: requireHandle(source.slides[0].handle),
        offsetX: asEmu(914400),
        offsetY: asEmu(457200),
        width: asEmu(2743200),
        height: asEmu(914400),
        text: "Initial textbox",
        name: "Added Textbox",
      }),
    );
    const addedShape = requireShape(findShapeByName(withTextBox, "Added Textbox"));
    const runHandle = addedShape.textBody?.paragraphs[0]?.runs[0]?.handle;
    if (runHandle === undefined || addedShape.handle === undefined) {
      throw new Error("added text box handles not found");
    }

    expectApplied(
      session.apply({
        kind: "replaceTextRunPlainText",
        handle: runHandle,
        text: "Edited textbox",
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "setShapeTransform",
        handle: addedShape.handle,
        offsetX: asEmu(1000),
        offsetY: asEmu(2000),
        width: asEmu(3000),
        height: asEmu(4000),
      }),
    );
    const rereadAdded = requireShape(findShapeByName(readPptx(writePptx(edited)), "Added Textbox"));

    expect(rereadAdded.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Edited textbox");
    expect(rereadAdded.transform).toMatchObject({
      offsetX: 1000,
      offsetY: 2000,
      width: 3000,
      height: 4000,
    });
  });

  it("adds a connector command and persists native connection sites", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const start = firstShape(source);
    const end = shapeWithoutTransform(source);
    const edited = expectApplied(
      session.apply({
        kind: "addConnector",
        slideHandle: requireHandle(source.slides[0].handle),
        preset: "curvedConnector3",
        offsetX: asEmu(100),
        offsetY: asEmu(200),
        width: asEmu(300),
        height: asEmu(400),
        start: {
          shapeHandle: requireHandle(start.handle),
          connectionSiteIndex: 0,
        },
        end: {
          shapeHandle: requireHandle(end.handle),
          connectionSiteIndex: 2,
        },
        outline: {
          tailEnd: { type: "triangle", width: "med", length: "med" },
        },
      }),
    );
    const connector = findConnectorByName(readPptx(writePptx(edited)), "Connector 12");

    expect(connector).toMatchObject({
      connection: {
        start: { shapeId: "10", connectionSiteIndex: 0 },
        end: { shapeId: "11", connectionSiteIndex: 2 },
      },
      geometry: { preset: "curvedConnector3" },
      outline: { tailEnd: { type: "triangle" } },
    });
  });

  it("undoes and redoes shape deletion for generated command sequences", async () => {
    const cases = [
      [{ kind: "deleteShape" }],
      [{ kind: "moveShape", offsetX: asEmu(1000), offsetY: asEmu(2000) }, { kind: "deleteShape" }],
      [{ kind: "resizeShape", width: asEmu(3000), height: asEmu(4000) }, { kind: "deleteShape" }],
    ] as const;

    for (const commands of cases) {
      const source = readPptx(await buildTextEditFixture());
      const session = createEditorSession(source);
      const handle = requireHandle(firstShape(source).handle);

      for (const command of commands) {
        expectApplied(session.apply({ ...command, handle }));
      }
      expect(findShapeNodeBySourceHandle(session.document, handle)).toBeUndefined();
      expect(firstRun(readPptx(writePptx(session.document))).text).toBe("No xfrm");

      for (let i = 0; i < commands.length; i += 1) expectHistory(session.undo());
      expect(findShapeNodeBySourceHandle(session.document, handle)).toBeDefined();
      expect(firstRun(readPptx(writePptx(session.document))).text).toBe("Original");

      for (let i = 0; i < commands.length; i += 1) expectHistory(session.redo());
      expect(findShapeNodeBySourceHandle(session.document, handle)).toBeUndefined();
      expect(firstRun(readPptx(writePptx(session.document))).text).toBe("No xfrm");
    }
  });

  it("rejects invalid add/delete shape commands without changing document state", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const missingHandle = {
      partPath: asPartPath("ppt/slides/slide1.xml"),
      nodeId: asSourceNodeId("999"),
      orderingSlot: 99,
    } satisfies SourceHandle;

    const invalidAdd = session.apply({
      kind: "addTextBox",
      slideHandle: requireHandle(source.slides[0].handle),
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(0),
      height: asEmu(100),
      text: "Invalid",
    });
    const missingDelete = session.apply({ kind: "deleteShape", handle: missingHandle });

    expect(invalidAdd).toMatchObject({ ok: false, code: "invalid-command" });
    expect(missingDelete).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(0);
  });
});

describe("EditorSession xfrm commands", () => {
  it("applies move and resize edits and persists them through write/read round-trip", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);

    expectApplied(
      session.apply({
        kind: "moveShape",
        handle,
        offsetX: asEmu(914400),
        offsetY: asEmu(1828800),
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "resizeShape",
        handle,
        width: asEmu(2743200),
        height: asEmu(914400),
      }),
    );
    const reread = readPptx(writePptx(edited));
    const rereadShape = requireShape(findShapeNodeBySourceHandle(reread, handle));

    expect(firstShape(source).transform).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 300,
      height: 400,
    });
    expect(requireShape(findShapeNodeBySourceHandle(edited, handle)).transform).toMatchObject({
      offsetX: 914400,
      offsetY: 1828800,
      width: 2743200,
      height: 914400,
    });
    expect(rereadShape.transform).toMatchObject({
      offsetX: 914400,
      offsetY: 1828800,
      width: 2743200,
      height: 914400,
    });
    expect(edited.edits?.filter((edit) => edit.kind === "updateShapeTransform")).toHaveLength(1);
  });

  it("undoes and redoes a move edit", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);

    expectApplied(
      session.apply({
        kind: "moveShape",
        handle,
        offsetX: asEmu(1000),
        offsetY: asEmu(2000),
      }),
    );

    const undone = expectHistory(session.undo());
    const redone = expectHistory(session.redo());

    expect(requireShape(findShapeNodeBySourceHandle(undone, handle)).transform).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 300,
      height: 400,
    });
    expect(
      requireShape(findShapeNodeBySourceHandle(readPptx(writePptx(undone)), handle)).transform,
    ).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 300,
      height: 400,
    });
    expect(requireShape(findShapeNodeBySourceHandle(redone, handle)).transform).toMatchObject({
      offsetX: 1000,
      offsetY: 2000,
      width: 300,
      height: 400,
    });
    expect(
      requireShape(findShapeNodeBySourceHandle(readPptx(writePptx(redone)), handle)).transform,
    ).toMatchObject({
      offsetX: 1000,
      offsetY: 2000,
      width: 300,
      height: 400,
    });
  });

  it("applies a full transform edit as one undoable command", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstShape(source).handle);

    const edited = expectApplied(
      session.apply({
        kind: "setShapeTransform",
        handle,
        offsetX: asEmu(1000),
        offsetY: asEmu(2000),
        width: asEmu(3000),
        height: asEmu(4000),
      }),
    );

    expect(session.undoDepth).toBe(1);
    expect(requireShape(findShapeNodeBySourceHandle(edited, handle)).transform).toMatchObject({
      offsetX: 1000,
      offsetY: 2000,
      width: 3000,
      height: 4000,
    });
    expect(edited.edits?.filter((edit) => edit.kind === "updateShapeTransform")).toHaveLength(1);

    expectHistory(session.undo());
    expect(firstShape(session.document).transform).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 300,
      height: 400,
    });
  });

  it("rejects invalid xfrm commands without changing document state or undo history", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const noXfrmHandle = requireHandle(shapeWithoutTransform(source).handle);
    const missingHandle = {
      partPath: asPartPath("ppt/slides/slide1.xml"),
      nodeId: asSourceNodeId("999"),
      orderingSlot: 99,
    } satisfies SourceHandle;

    const noXfrmResult = session.apply({
      kind: "moveShape",
      handle: noXfrmHandle,
      offsetX: asEmu(1000),
      offsetY: asEmu(2000),
    });
    const missingHandleResult = session.apply({
      kind: "resizeShape",
      handle: missingHandle,
      width: asEmu(3000),
      height: asEmu(4000),
    });
    const invalidExtentResult = session.apply({
      kind: "resizeShape",
      handle: requireHandle(firstShape(source).handle),
      width: asEmu(0),
      height: asEmu(Number.NaN),
    });

    expect(noXfrmResult).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(noXfrmResult.ok).toBe(false);
    if (!noXfrmResult.ok) {
      expect(noXfrmResult.message).toMatch(/does not reference a shape with xfrm/);
    }
    expect(missingHandleResult).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(missingHandleResult.ok).toBe(false);
    if (!missingHandleResult.ok) {
      expect(missingHandleResult.message).toMatch(/shape handle was not found/);
    }
    expect(invalidExtentResult).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(invalidExtentResult.ok).toBe(false);
    if (!invalidExtentResult.ok) {
      expect(invalidExtentResult.message).toMatch(/finite positive EMU value/);
    }
    expect(session.document).toBe(before);
    expect(firstShape(session.document).transform).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 300,
      height: 400,
    });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });
});

describe("EditorSession shape style commands", () => {
  it("sets shape fill and shape or connector outlines and persists them through write/read", async () => {
    const source = readPptx(await buildShapeStyleFixture());
    const session = createEditorSession(source);
    const shapeHandle = requireHandle(firstShape(source).handle);
    const connectorHandle = requireHandle(findConnectorByName(source, "Connector").handle);

    expectApplied(
      session.apply({
        kind: "setShapeFill",
        handle: shapeHandle,
        fill: { kind: "solid", color: { kind: "srgb", hex: "00aa44" } },
      }),
    );
    expectApplied(
      session.apply({
        kind: "setShapeOutline",
        handle: shapeHandle,
        outline: {
          width: asEmu(25400),
          fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
        },
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "setShapeOutline",
        handle: connectorHandle,
        outline: {
          width: asEmu(38100),
          fill: { kind: "solid", color: { kind: "srgb", hex: "AA00FF" } },
        },
      }),
    );
    const reread = readPptx(writePptx(edited));

    expect(firstShape(reread).fill).toEqual({
      kind: "solid",
      color: { kind: "srgb", hex: "00AA44" },
    });
    expect(firstShape(reread).outline).toMatchObject({
      width: 25400,
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
    });
    expect(findConnectorByName(reread, "Connector").outline).toMatchObject({
      width: 38100,
      fill: { kind: "solid", color: { kind: "srgb", hex: "AA00FF" } },
    });
  });

  it("undoes and redoes shape fill and noFill outline edits", async () => {
    const source = readPptx(await buildShapeStyleFixture());
    const session = createEditorSession(source);
    const shapeHandle = requireHandle(firstShape(source).handle);

    expectApplied(
      session.applyAll([
        { kind: "setShapeFill", handle: shapeHandle, fill: { kind: "none" } },
        {
          kind: "setShapeOutline",
          handle: shapeHandle,
          outline: { fill: { kind: "none" } },
        },
      ]),
    );

    const undone = expectHistory(session.undo());
    const redone = expectHistory(session.redo());

    expect(firstShape(undone).fill).toBeUndefined();
    expect(firstShape(readPptx(writePptx(undone))).fill).toBeUndefined();
    expect(firstShape(redone).fill).toEqual({ kind: "none" });
    expect(firstShape(redone).outline).toMatchObject({ fill: { kind: "none" } });
    expect(firstShape(readPptx(writePptx(redone))).fill).toEqual({ kind: "none" });
    expect(firstShape(readPptx(writePptx(redone))).outline).toMatchObject({
      fill: { kind: "none" },
    });
  });

  it("keeps only the latest generated shape style edit per target", async () => {
    const source = readPptx(await buildShapeStyleFixture());
    const session = createEditorSession(source);
    const shapeHandle = requireHandle(firstShape(source).handle);

    expectApplied(
      session.apply({
        kind: "setShapeFill",
        handle: shapeHandle,
        fill: { kind: "solid", color: { kind: "srgb", hex: "111111" } },
      }),
    );
    expectApplied(
      session.apply({
        kind: "setShapeFill",
        handle: shapeHandle,
        fill: { kind: "solid", color: { kind: "srgb", hex: "222222" } },
      }),
    );
    expectApplied(
      session.apply({
        kind: "setShapeOutline",
        handle: shapeHandle,
        outline: { fill: { kind: "solid", color: { kind: "srgb", hex: "445566" } } },
      }),
    );
    expectApplied(
      session.apply({
        kind: "setShapeOutline",
        handle: shapeHandle,
        outline: { width: asEmu(12700) },
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "setShapeOutline",
        handle: shapeHandle,
        outline: { width: asEmu(25400) },
      }),
    );

    expect(edited.edits?.filter((edit) => edit.kind === "updateShapeFill")).toEqual([
      {
        kind: "updateShapeFill",
        handle: shapeHandle,
        fill: { kind: "solid", color: { kind: "srgb", hex: "222222" } },
      },
    ]);
    expect(edited.edits?.filter((edit) => edit.kind === "updateShapeOutline")).toEqual([
      {
        kind: "updateShapeOutline",
        handle: shapeHandle,
        outline: {
          fill: { kind: "solid", color: { kind: "srgb", hex: "445566" } },
          width: 25400,
        },
      },
    ]);
    expect(firstShape(readPptx(writePptx(edited))).outline).toMatchObject({
      fill: { kind: "solid", color: { kind: "srgb", hex: "445566" } },
      width: 25400,
    });
  });

  it("rejects invalid shape style commands without changing document state or undo history", async () => {
    const source = readPptx(await buildShapeStyleFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const shapeHandle = requireHandle(firstShape(source).handle);
    const connectorHandle = requireHandle(findConnectorByName(source, "Connector").handle);

    const invalidFill = session.apply({
      kind: "setShapeFill",
      handle: shapeHandle,
      fill: { kind: "solid", color: { kind: "srgb", hex: "bad" } },
    });
    const invalidOutline = session.apply({
      kind: "setShapeOutline",
      handle: shapeHandle,
      outline: { width: asEmu(0) },
    });
    const connectorFill = session.apply({
      kind: "setShapeFill",
      handle: connectorHandle,
      fill: { kind: "none" },
    });

    for (const result of [invalidFill, invalidOutline, connectorFill]) {
      expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    }
    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });
});
