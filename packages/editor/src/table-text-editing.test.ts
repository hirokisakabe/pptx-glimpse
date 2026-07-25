import {
  addTable,
  addTextBox,
  asEmu,
  asSourceNodeId,
  createPptx,
  readPptx,
  type SourceHandle,
  type SourceTable,
  type SourceTextRun,
  writePptx,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";

describe("EditorSession existing table cell text", () => {
  it("uses text convenience methods and commands with undo and redo", () => {
    const source = buildExistingTableSource();
    const table = firstTable(source);
    const run = table.table.rows[0].cells[0].textBody!.paragraphs[0].runs[0];
    const paragraph = table.table.rows[0].cells[1].textBody!.paragraphs[0];
    const session = createEditorSession(source);

    expect(session.replaceTextRunPlainText(run, "Edited through convenience")).toMatchObject({
      ok: true,
    });
    expect(tableRunText(session.document, 0, 0)).toBe("Edited through convenience");
    expect(session.undoDepth).toBe(1);
    expect(session.undo()).toMatchObject({ ok: true });
    expect(tableRunText(session.document, 0, 0)).toBe("Run target");
    expect(session.redo()).toMatchObject({ ok: true });
    expect(tableRunText(session.document, 0, 0)).toBe("Edited through convenience");

    const result = session.applyAll([
      {
        kind: "setTextRunProperties",
        handle: requireHandle(run.handle),
        properties: { bold: true },
      },
      {
        kind: "replaceParagraphPlainText",
        handle: requireHandle(paragraph.handle),
        text: "Edited through command",
      },
      {
        kind: "setParagraphProperties",
        handle: requireHandle(
          firstTable(session.document).table.rows[1].cells[1].textBody!.paragraphs[0].handle,
        ),
        properties: { align: "right" },
      },
    ]);

    expect(result).toMatchObject({ ok: true });
    expect(
      firstTable(session.document).table.rows[0].cells[0].textBody?.paragraphs[0].runs[0]
        .properties,
    ).toMatchObject({ bold: true });
    expect(tableRunText(session.document, 0, 1)).toBe("Edited through command");
    expect(
      firstTable(session.document).table.rows[1].cells[1].textBody?.paragraphs[0].properties,
    ).toMatchObject({ align: "right" });
    expect(session.undoDepth).toBe(2);
    expect(session.undo()).toMatchObject({ ok: true });
    expect(tableRunText(session.document, 0, 1)).toBe("Paragraph target");
    expect(session.redo()).toMatchObject({ ok: true });
    expect(tableRunText(session.document, 0, 1)).toBe("Edited through command");
  });

  it("rejects handleless, foreign, and absent targets atomically", () => {
    const source = buildExistingTableSource();
    const foreign = buildExistingTableSource(true);
    const run = firstTable(source).table.rows[0].cells[0].textBody!.paragraphs[0].runs[0];
    const foreignRun = firstTable(foreign).table.rows[0].cells[0].textBody!.paragraphs[0].runs[0];
    const session = createEditorSession(source);
    const originalDocument = session.document;
    const withoutHandle: SourceTextRun = { kind: "textRun", text: "No handle" };

    expect(session.replaceTextRunPlainText(withoutHandle, "Rejected")).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(session.replaceTextRunPlainText(foreignRun, "Rejected")).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(
      session.apply({
        kind: "replaceTextRunPlainText",
        handle: {
          ...requireHandle(run.handle),
          nodeId: asSourceNodeId("text:table:999:row:0:cell:0:p:0:r:0"),
        },
        text: "Rejected",
      }),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(
      session.applyAll([
        {
          kind: "replaceParagraphPlainText",
          handle: requireHandle(
            firstTable(source).table.rows[0].cells[0].textBody!.paragraphs[0].handle,
          ),
          text: "Conflicting paragraph edit",
        },
        {
          kind: "setTextRunProperties",
          handle: requireHandle(run.handle),
          properties: { bold: true },
        },
      ]),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(originalDocument);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
    expect(tableRunText(session.document, 0, 0)).toBe("Run target");
  });
});

function buildExistingTableSource(withLeadingShape = false) {
  const source = createPptx();
  const withOptionalShape = withLeadingShape
    ? addTextBox(source, source.slides[0].handle!, {
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(100),
        height: asEmu(100),
        text: "Leading",
      })
    : source;
  const authored = addTable(withOptionalShape, withOptionalShape.slides[0].handle!, {
    offsetX: asEmu(100),
    offsetY: asEmu(100),
    width: asEmu(4000),
    height: asEmu(2000),
    columnWidths: [asEmu(2000), asEmu(2000)],
    rows: [
      {
        height: asEmu(1000),
        cells: [{ text: "Run target" }, { text: "Paragraph target" }],
      },
      {
        height: asEmu(1000),
        cells: [{ text: "Run properties" }, { text: "Paragraph properties" }],
      },
    ],
  });
  return readPptx(writePptx(authored));
}

function firstTable(source: ReturnType<typeof buildExistingTableSource>): SourceTable {
  const table = source.slides[0].shapes.find((shape) => shape.kind === "table");
  if (table === undefined) throw new Error("test fixture table is missing");
  return table;
}

function tableRunText(
  source: ReturnType<typeof buildExistingTableSource>,
  rowIndex: number,
  cellIndex: number,
): string | undefined {
  return firstTable(source).table.rows[rowIndex].cells[cellIndex].textBody?.paragraphs[0].runs[0]
    .text;
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("test fixture handle is missing");
  return handle;
}
