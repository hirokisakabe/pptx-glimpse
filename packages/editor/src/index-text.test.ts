import {
  asPartPath,
  asPt,
  asSourceNodeId,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  writePptx,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";
import {
  buildTextEditFixture,
  expectApplied,
  expectHistory,
  firstParagraph,
  firstRun,
  firstShape,
  requireHandle,
} from "./index.test-helpers.js";

describe("EditorSession text-run commands", () => {
  it("applies a text-run edit and persists it through write/read round-trip", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const run = firstRun(source);

    const edited = expectApplied(
      session.apply({
        kind: "replaceTextRunPlainText",
        handle: requireHandle(run.handle),
        text: "Edited text",
      }),
    );
    const reread = readPptx(writePptx(edited));

    expect(firstRun(source).text).toBe("Original");
    expect(firstRun(session.document).text).toBe("Edited text");
    expect(firstRun(reread).text).toBe("Edited text");
    expect(firstParagraph(reread).runs[1].text).toBe(" Keep ");
  });

  it("undoes and redoes a text-run edit", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);

    expectApplied(
      session.apply({
        kind: "replaceTextRunPlainText",
        handle: requireHandle(firstRun(source).handle),
        text: "Edited text",
      }),
    );

    const undone = expectHistory(session.undo());
    const redone = expectHistory(session.redo());

    expect(firstRun(undone).text).toBe("Original");
    expect(firstRun(readPptx(writePptx(undone))).text).toBe("Original");
    expect(firstRun(redone).text).toBe("Edited text");
    expect(firstRun(readPptx(writePptx(redone))).text).toBe("Edited text");
    expect(session.canUndo).toBe(true);
    expect(session.canRedo).toBe(false);
  });

  it("returns common failures for empty history without changing document or selection", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const shapeHandle = requireHandle(firstShape(source).handle);
    expect(session.selectShape(shapeHandle)).toMatchObject({ ok: true });

    expect(session.undo()).toEqual({
      ok: false,
      code: "empty-undo-stack",
      message: "undo: undo history is empty",
    });
    expect(session.redo()).toEqual({
      ok: false,
      code: "empty-redo-stack",
      message: "redo: redo history is empty",
    });
    expect(session.document).toBe(source);
    expect(session.selection).toEqual({ shapeHandle });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("keeps redo history when applying a no-op command after undo", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstRun(source).handle);

    expectApplied(session.apply({ kind: "replaceTextRunPlainText", handle, text: "Edited text" }));
    expectHistory(session.undo());
    const noOp = expectApplied(
      session.apply({ kind: "replaceTextRunPlainText", handle, text: "Original" }),
    );

    expect(noOp).toBe(source);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(1);
    expectHistory(session.redo());
    expect(firstRun(session.document).text).toBe("Edited text");
  });

  it("keeps the latest edit when the same text run is edited repeatedly", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstRun(source).handle);

    expectApplied(
      session.apply({
        kind: "replaceTextRunPlainText",
        handle,
        text: "First edit",
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "replaceTextRunPlainText",
        handle,
        text: "Second edit",
      }),
    );
    const reread = readPptx(writePptx(edited));

    expect(firstRun(edited).text).toBe("Second edit");
    expect(firstRun(reread).text).toBe("Second edit");
  });

  it("rejects an invalid command without changing document state or undo history", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const invalidHandle = {
      partPath: asPartPath("ppt/slides/slide1.xml"),
      nodeId: asSourceNodeId("text:shape:999:p:0:r:0"),
      orderingSlot: 0,
    } satisfies SourceHandle;

    const result = session.apply({
      kind: "replaceTextRunPlainText",
      handle: invalidHandle,
      text: "Should not apply",
    });

    expect(result).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.message).toMatch(/text run handle was not found/);
    }
    expect(session.document).toBe(before);
    expect(firstRun(session.document).text).toBe("Original");
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
    expect(session.canUndo).toBe(false);
    expect(session.canRedo).toBe(false);
    expect(session.undo()).toEqual({
      ok: false,
      code: "empty-undo-stack",
      message: "undo: undo history is empty",
    });
  });

  it("rejects an invalid command batch without partially applying earlier commands", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const validHandle = requireHandle(firstRun(source).handle);
    const invalidHandle = {
      partPath: asPartPath("ppt/slides/slide1.xml"),
      nodeId: asSourceNodeId("text:shape:999:p:0:r:0"),
      orderingSlot: 0,
    } satisfies SourceHandle;

    const result = session.applyAll([
      {
        kind: "replaceTextRunPlainText",
        handle: validHandle,
        text: "Should not stay applied",
      },
      {
        kind: "replaceTextRunPlainText",
        handle: invalidHandle,
        text: "Should not apply",
      },
    ]);

    expect(result).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(session.document).toBe(before);
    expect(firstRun(session.document).text).toBe("Original");
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("does not convert an unsupported command kind into an expected failure", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);

    expect(() =>
      session.apply({
        // @ts-expect-error exercises a programmer error from an untyped JavaScript caller.
        kind: "unsupported-command",
      }),
    ).toThrow(TypeError);
    expect(session.document).toBe(source);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("does not convert an unexpected failure from a supported command into a rejection", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const unexpected = new Error("unexpected command implementation failure");

    expect(() =>
      session.apply({
        kind: "replaceTextRunPlainText",
        handle: requireHandle(firstRun(source).handle),
        get text(): string {
          throw unexpected;
        },
      }),
    ).toThrow(unexpected);
    expect(session.document).toBe(source);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("rejects non-string text from JavaScript callers", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);

    const runResult = session.apply({
      kind: "replaceTextRunPlainText",
      handle: requireHandle(firstRun(source).handle),
      // @ts-expect-error exercises runtime validation for JavaScript callers.
      text: 42,
    });
    const paragraphResult = session.apply({
      kind: "replaceParagraphPlainText",
      handle: requireHandle(firstParagraph(source).handle),
      // @ts-expect-error exercises runtime validation for JavaScript callers.
      text: null,
    });

    expect(runResult).toMatchObject({ ok: false, code: "invalid-command" });
    expect(paragraphResult).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(source);
    expect(session.undoDepth).toBe(0);
  });
});

describe("EditorSession text run property commands", () => {
  it("sets and clears supported run properties and persists them through write/read", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstRun(source).handle);

    const setDocument = expectApplied(
      session.apply({
        kind: "setTextRunProperties",
        handle,
        properties: {
          bold: false,
          italic: false,
          underline: true,
          fontSize: asPt(30),
          color: { kind: "srgb", hex: "336699" },
          typeface: "Liberation Sans",
        },
      }),
    );
    const clearedDocument = expectApplied(
      session.apply({
        kind: "clearTextRunProperties",
        handle,
        properties: ["italic", "fontSize", "color"],
      }),
    );
    const reread = readPptx(writePptx(clearedDocument));

    expect(firstRun(setDocument).properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 30,
      color: { kind: "srgb", hex: "336699" },
      typeface: "Liberation Sans",
    });
    expect(firstRun(reread).properties).toMatchObject({
      bold: false,
      underline: true,
      typeface: "Liberation Sans",
    });
    expect(firstRun(reread).properties?.italic).toBeUndefined();
    expect(firstRun(reread).properties?.fontSize).toBeUndefined();
    expect(firstRun(reread).properties?.color).toBeUndefined();
    expect(firstParagraph(reread).runs[1].properties).toMatchObject({
      italic: true,
      fontSize: 18,
      typeface: "Arial",
    });
  });

  it("undoes and redoes run property edits", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstRun(source).handle);

    expectApplied(
      session.apply({
        kind: "setTextRunProperties",
        handle,
        properties: { underline: true, color: { kind: "srgb", hex: "0088CC" } },
      }),
    );

    const undone = expectHistory(session.undo());
    const redone = expectHistory(session.redo());

    expect(firstRun(undone).properties?.underline).toBeUndefined();
    expect(firstRun(readPptx(writePptx(undone))).properties?.underline).toBeUndefined();
    expect(firstRun(redone).properties).toMatchObject({
      underline: true,
      color: { kind: "srgb", hex: "0088CC" },
    });
    expect(firstRun(readPptx(writePptx(redone))).properties).toMatchObject({
      underline: true,
      color: { kind: "srgb", hex: "0088CC" },
    });
  });

  it("rejects invalid run property commands without changing document state or undo history", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const handle = requireHandle(firstRun(source).handle);
    const invalidHandle = {
      partPath: asPartPath("ppt/slides/slide1.xml"),
      nodeId: asSourceNodeId("text:shape:999:p:0:r:0"),
      orderingSlot: 0,
    } satisfies SourceHandle;

    const invalidHex = session.apply({
      kind: "setTextRunProperties",
      handle,
      properties: { color: { kind: "srgb", hex: "bad" } },
    });
    const invalidFontSize = session.apply({
      kind: "setTextRunProperties",
      handle,
      properties: { fontSize: asPt(0) },
    });
    const unsupportedSetProperty = session.apply({
      kind: "setTextRunProperties",
      handle,
      // @ts-expect-error exercises runtime validation for JS callers.
      properties: { strikethrough: true },
    });
    const emptyClearProperties = session.apply({
      kind: "clearTextRunProperties",
      handle,
      properties: [],
    });
    const missingHandle = session.apply({
      kind: "setTextRunProperties",
      handle: invalidHandle,
      properties: { bold: true },
    });

    for (const result of [
      invalidHex,
      invalidFontSize,
      unsupportedSetProperty,
      emptyClearProperties,
      missingHandle,
    ]) {
      expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    }
    expect(session.document).toBe(before);
    expect(firstRun(session.document).properties).toEqual(firstRun(source).properties);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("keeps only the latest generated edit per run property while preserving independent properties", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstRun(source).handle);

    expectApplied(
      session.apply({
        kind: "setTextRunProperties",
        handle,
        properties: { bold: false, italic: false, fontSize: asPt(20) },
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "setTextRunProperties",
        handle,
        properties: { bold: true, color: { kind: "srgb", hex: "445566" } },
      }),
    );
    const propertyEdits =
      edited.edits?.filter((edit) => edit.kind === "updateTextRunProperties") ?? [];

    expect(propertyEdits).toHaveLength(2);
    expect(propertyEdits[0]).toMatchObject({ set: { italic: false, fontSize: 20 } });
    expect(propertyEdits[1]).toMatchObject({
      set: { bold: true, color: { kind: "srgb", hex: "445566" } },
    });
    expect(firstRun(readPptx(writePptx(edited))).properties).toMatchObject({
      bold: true,
      italic: false,
      fontSize: 20,
      color: { kind: "srgb", hex: "445566" },
    });
  });

  it("preserves mixed clear and set edits from an existing edit journal during normalization", async () => {
    const source = readPptx(await buildTextEditFixture());
    const handle = requireHandle(firstRun(source).handle);
    const sourceWithMixedEdit: PptxSourceModel = {
      ...source,
      edits: [
        {
          kind: "updateTextRunProperties",
          handle,
          clear: ["color"],
          set: { color: { kind: "srgb", hex: "112233" } },
        },
      ],
    };
    const session = createEditorSession(sourceWithMixedEdit);
    const edited = expectApplied(
      session.apply({
        kind: "setTextRunProperties",
        handle,
        properties: { bold: false },
      }),
    );
    const propertyEdits =
      edited.edits?.filter((edit) => edit.kind === "updateTextRunProperties") ?? [];
    const reread = readPptx(writePptx(edited));

    expect(propertyEdits).toHaveLength(2);
    expect(propertyEdits[0]).toMatchObject({
      set: { color: { kind: "srgb", hex: "112233" } },
    });
    expect(propertyEdits[0]?.clear).toBeUndefined();
    expect(propertyEdits[1]).toMatchObject({ set: { bold: false } });
    expect(firstRun(reread).properties).toMatchObject({
      bold: false,
      color: { kind: "srgb", hex: "112233" },
    });
  });

  it("passes property-style undo and redo checks for generated decoration command sequences", async () => {
    const cases = [
      [
        { kind: "setTextRunProperties", properties: { bold: false } },
        { kind: "setTextRunProperties", properties: { italic: true } },
      ],
      [
        { kind: "setTextRunProperties", properties: { underline: true } },
        { kind: "clearTextRunProperties", properties: ["underline"] },
      ],
      [
        { kind: "clearTextRunProperties", properties: ["fontSize", "typeface"] },
        { kind: "setTextRunProperties", properties: { fontSize: asPt(22), typeface: "Arial" } },
      ],
      [
        { kind: "setTextRunProperties", properties: { color: { kind: "srgb", hex: "123ABC" } } },
        { kind: "clearTextRunProperties", properties: ["color"] },
      ],
    ] as const;

    for (const commands of cases) {
      const source = readPptx(await buildTextEditFixture());
      const session = createEditorSession(source);
      const handle = requireHandle(firstRun(source).handle);
      for (const command of commands) {
        expectApplied(session.apply({ ...command, handle }));
      }
      const edited = session.document;

      for (let i = 0; i < commands.length; i += 1) expectHistory(session.undo());
      expect(firstRun(session.document).properties).toEqual(firstRun(source).properties);
      expect(firstRun(readPptx(writePptx(session.document))).properties).toEqual(
        firstRun(source).properties,
      );

      for (let i = 0; i < commands.length; i += 1) expectHistory(session.redo());
      expect(firstRun(session.document).properties).toEqual(firstRun(edited).properties);
      expect(firstRun(readPptx(writePptx(session.document))).properties).toEqual(
        firstRun(readPptx(writePptx(edited))).properties,
      );
    }
  });
});

describe("EditorSession paragraph property commands", () => {
  it("replaces paragraph text as one undoable writer-persisted command", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstParagraph(source).handle);

    const edited = expectApplied(
      session.apply({
        kind: "replaceParagraphPlainText",
        handle,
        text: "Paragraph replacement",
      }),
    );

    expect(firstParagraph(edited).runs.map((run) => run.text)).toEqual(["Paragraph replacement"]);
    expect(firstParagraph(readPptx(writePptx(edited))).runs.map((run) => run.text)).toEqual([
      "Paragraph replacement",
    ]);
    expect(session.undoDepth).toBe(1);
    expect(firstParagraph(expectHistory(session.undo())).runs.map((run) => run.text)).toEqual([
      "Original",
      " Keep ",
    ]);
    expect(firstParagraph(expectHistory(session.redo())).runs.map((run) => run.text)).toEqual([
      "Paragraph replacement",
    ]);
  });

  it("normalizes an earlier run text edit into paragraph replacement and rejects the reverse order", async () => {
    const source = readPptx(await buildTextEditFixture());
    const paragraphHandle = requireHandle(firstParagraph(source).handle);
    const runHandle = requireHandle(firstRun(source).handle);
    const runThenParagraph = createEditorSession(source);

    const edited = expectApplied(
      runThenParagraph.applyAll([
        { kind: "replaceTextRunPlainText", handle: runHandle, text: "Intermediate" },
        {
          kind: "replaceParagraphPlainText",
          handle: paragraphHandle,
          text: "Final paragraph",
        },
      ]),
    );

    expect(firstParagraph(edited).runs.map((run) => run.text)).toEqual(["Final paragraph"]);
    expect(firstParagraph(readPptx(writePptx(edited))).runs.map((run) => run.text)).toEqual([
      "Final paragraph",
    ]);
    expect(runThenParagraph.undoDepth).toBe(1);

    const paragraphThenRun = createEditorSession(source);
    expect(
      paragraphThenRun.applyAll([
        {
          kind: "replaceParagraphPlainText",
          handle: paragraphHandle,
          text: "Intermediate paragraph",
        },
        { kind: "replaceTextRunPlainText", handle: runHandle, text: "Unsafe follow-up" },
      ]),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(paragraphThenRun.document).toBe(source);
    expect(paragraphThenRun.undoDepth).toBe(0);
  });

  it("rejects paragraph replacement combined with run edits for the same paragraph", async () => {
    const source = readPptx(await buildTextEditFixture());
    const paragraphHandle = requireHandle(firstParagraph(source).handle);
    const runHandle = requireHandle(firstRun(source).handle);

    for (const [orderIndex, commands] of (
      [
        [
          {
            kind: "setTextRunProperties",
            handle: runHandle,
            properties: { bold: false },
          },
          {
            kind: "replaceParagraphPlainText",
            handle: paragraphHandle,
            text: "Replacement",
          },
        ],
        [
          {
            kind: "replaceParagraphPlainText",
            handle: paragraphHandle,
            text: "Replacement",
          },
          {
            kind: "setTextRunProperties",
            handle: runHandle,
            properties: { bold: false },
          },
        ],
      ] as const
    ).entries()) {
      const session = createEditorSession(source);
      const result = session.applyAll(commands);

      expect(result, `command order ${String(orderIndex)}`).toMatchObject({
        ok: false,
        code: "invalid-command",
      });
      expect(session.document).toBe(source);
      expect(session.undoDepth).toBe(0);
    }

    const runThenParagraph = createEditorSession(source);
    expectApplied(
      runThenParagraph.apply({
        kind: "setTextRunProperties",
        handle: runHandle,
        properties: { bold: false },
      }),
    );
    expect(
      runThenParagraph.apply({
        kind: "replaceParagraphPlainText",
        handle: paragraphHandle,
        text: "Replacement",
      }),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(runThenParagraph.undoDepth).toBe(1);
    expect(firstRun(runThenParagraph.document).properties?.bold).toBe(false);

    const paragraphThenRun = createEditorSession(source);
    expectApplied(
      paragraphThenRun.apply({
        kind: "replaceParagraphPlainText",
        handle: paragraphHandle,
        text: "Replacement",
      }),
    );
    expect(
      paragraphThenRun.apply({
        kind: "setTextRunProperties",
        handle: runHandle,
        properties: { bold: false },
      }),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(paragraphThenRun.undoDepth).toBe(1);
    expect(firstParagraph(paragraphThenRun.document).runs[0].text).toBe("Replacement");
  });

  it("applies paragraph alignment and bullet edits and persists them through write/read", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstParagraph(source).handle);

    const edited = expectApplied(
      session.apply({
        kind: "setParagraphProperties",
        handle,
        properties: {
          align: "right",
          level: 2,
          bullet: { type: "char", char: "\u2022" },
        },
      }),
    );
    const reread = readPptx(writePptx(edited));

    expect(firstParagraph(session.document).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(firstParagraph(reread).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
  });

  it("undoes and redoes paragraph property edits", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstParagraph(source).handle);

    expectApplied(
      session.apply({
        kind: "setParagraphProperties",
        handle,
        properties: { align: "justify", bullet: { type: "none" } },
      }),
    );

    const undone = expectHistory(session.undo());
    const redone = expectHistory(session.redo());

    expect(firstParagraph(undone).properties).toEqual(firstParagraph(source).properties);
    expect(firstParagraph(readPptx(writePptx(undone))).properties).toEqual(
      firstParagraph(source).properties,
    );
    expect(firstParagraph(redone).properties).toMatchObject({
      align: "justify",
      bullet: { type: "none" },
    });
    expect(firstParagraph(readPptx(writePptx(redone))).properties).toMatchObject({
      align: "justify",
      bullet: { type: "none" },
    });
  });

  it("keeps only the latest paragraph property edits and skips same-value repeats", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const handle = requireHandle(firstParagraph(source).handle);

    expectApplied(
      session.apply({
        kind: "setParagraphProperties",
        handle,
        properties: { align: "left", level: 1, bullet: { type: "char", char: "\u2022" } },
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "setParagraphProperties",
        handle,
        properties: { align: "right", bullet: { type: "none" } },
      }),
    );
    const repeated = expectApplied(
      session.apply({
        kind: "setParagraphProperties",
        handle,
        properties: { align: "right", bullet: { type: "none" } },
      }),
    );
    const propertyEdits =
      repeated.edits?.filter((edit) => edit.kind === "updateParagraphProperties") ?? [];

    expect(repeated).toBe(edited);
    expect(propertyEdits).toHaveLength(2);
    expect(propertyEdits[0]).toMatchObject({ set: { level: 1 } });
    expect(propertyEdits[1]).toMatchObject({
      set: { align: "right", bullet: { type: "none" } },
    });
  });

  it("rejects invalid paragraph property commands without changing history", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const handle = requireHandle(firstParagraph(source).handle);

    const invalidAlign = session.apply({
      kind: "setParagraphProperties",
      handle,
      // @ts-expect-error exercises runtime validation for JS callers.
      properties: { align: "middle" },
    });
    const invalidClear = session.apply({
      kind: "clearParagraphProperties",
      handle,
      // @ts-expect-error exercises runtime validation for JS callers.
      properties: ["spacing"],
    });

    expect(invalidAlign).toMatchObject({ ok: false, code: "invalid-command" });
    expect(invalidClear).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(0);
  });
});
