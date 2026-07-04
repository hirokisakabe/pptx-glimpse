import {
  asEmu,
  asPartPath,
  asSourceNodeId,
  findShapeNodeBySourceHandle,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  type SourceShape,
  type SourceShapeNode,
  writePptx,
} from "@pptx-glimpse/document";
import JSZip from "jszip";
import { Node as ProseMirrorNode } from "prosemirror-model";
import { Transform } from "prosemirror-transform";
import { describe, expect, it } from "vitest";

import {
  createEditorSession,
  type EditorApplyCommandResult,
  type EditorHistoryResult,
  pptxTextBodySchema,
  proseMirrorDocJsonToEditorCommands,
  proseMirrorDocJsonToTextBody,
  textBodyToProseMirrorDocJson,
} from "./index.js";

const encoder = new TextEncoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

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
    expect(session.undo()).toEqual({ ok: false, reason: "empty-undo-stack" });
  });
});

describe("ProseMirror text body conversion", () => {
  it("round-trips paragraph and run text with source properties", async () => {
    const source = readPptx(await buildTextEditFixture());
    const textBody = firstTextBody(source);
    const docJson = textBodyToProseMirrorDocJson(textBody);
    const roundTripped = proseMirrorDocJsonToTextBody(textBody, docJson);

    expect(roundTripped.paragraphs).toEqual(textBody.paragraphs);
    expect(roundTripped.paragraphs[0].runs[0].properties).toEqual(
      textBody.paragraphs[0].runs[0].properties,
    );
    expect(roundTripped.paragraphs[0].runs[1].properties).toEqual(
      textBody.paragraphs[0].runs[1].properties,
    );
  });

  it("turns a run-crossing ProseMirror replacement into writer-persisted run edits", async () => {
    const source = readPptx(await buildTextEditFixture());
    const textBody = firstTextBody(source);
    const doc = ProseMirrorNode.fromJSON(
      pptxTextBodySchema,
      textBodyToProseMirrorDocJson(textBody),
    );
    const firstRunMark = doc.firstChild?.child(0).marks[0];
    if (firstRunMark === undefined) throw new Error("first run mark not found");
    const transform = new Transform(doc).replaceWith(
      1 + "Orig".length,
      1 + "Original Ke".length,
      pptxTextBodySchema.text("X", [firstRunMark]),
    );
    const editedJson: unknown = transform.doc.toJSON();
    const commands = proseMirrorDocJsonToEditorCommands(textBody, editedJson);
    const session = createEditorSession(source);

    for (const command of commands) {
      expectApplied(session.apply(command));
    }
    const reread = readPptx(writePptx(session.document));

    expect(commands).toHaveLength(2);
    expect(firstParagraph(session.document).runs.map((run) => run.text)).toEqual(["OrigX", "ep "]);
    expect(firstParagraph(reread).runs.map((run) => run.text)).toEqual(["OrigX", "ep "]);
    expect(firstParagraph(reread).runs[0].properties).toEqual(
      firstParagraph(source).runs[0].properties,
    );
    expect(firstParagraph(reread).runs[1].properties).toEqual(
      firstParagraph(source).runs[1].properties,
    );
  });

  it("falls back to paragraph replacement when run mark order changes", async () => {
    const source = readPptx(await buildTextEditFixture());
    const textBody = firstTextBody(source);
    const docJson = textBodyToProseMirrorDocJson(textBody);
    const paragraph = docJson.content?.[0];
    const firstText = paragraph?.content?.[0];
    const secondText = paragraph?.content?.[1];
    if (paragraph === undefined || firstText === undefined || secondText === undefined) {
      throw new Error("text body doc fixture is missing expected text nodes");
    }
    const editedJson = {
      type: "doc",
      content: [
        {
          ...paragraph,
          content: [secondText, firstText],
        },
      ],
    };
    const commands = proseMirrorDocJsonToEditorCommands(textBody, editedJson);
    const session = createEditorSession(source);

    for (const command of commands) {
      expectApplied(session.apply(command));
    }
    const reread = readPptx(writePptx(session.document));

    expect(commands).toEqual([
      {
        kind: "replaceParagraphPlainText",
        handle: firstParagraph(source).handle,
        text: " Keep Original",
      },
    ]);
    expect(firstParagraph(reread).runs.map((run) => run.text)).toEqual([" Keep Original"]);
  });

  it("rejects text bodies with empty or unsupported run-like content", async () => {
    const source = readPptx(await buildTextEditFixture());
    const textBody = firstTextBody(source);
    const paragraph = textBody.paragraphs[0];
    const run = paragraph.runs[0];
    const withEmptyRun = {
      ...textBody,
      paragraphs: [{ ...paragraph, runs: [{ ...run, text: "" }] }],
    };
    const withBreakRun = {
      ...textBody,
      paragraphs: [{ ...paragraph, runs: [{ ...run, text: "\n" }] }],
    };

    expect(() => textBodyToProseMirrorDocJson(withEmptyRun)).toThrow(/empty text runs/);
    expect(() => textBodyToProseMirrorDocJson(withBreakRun)).toThrow(/unsupported run-like/);
  });
});

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

async function buildTextEditFixture(): Promise<Uint8Array> {
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
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:pPr algn="ctr"/>` +
        `<a:r><a:rPr b="1" sz="2400"><a:latin typeface="Aptos"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Original</a:t></a:r>` +
        `<a:r><a:rPr i="1" sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t xml:space="preserve"> Keep </a:t></a:r>` +
        `</a:p>` +
        `</p:txBody></p:sp>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="11" name="No xfrm"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>No xfrm</a:t></a:r></a:p></p:txBody></p:sp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

function expectApplied(result: EditorApplyCommandResult): PptxSourceModel {
  if (!result.ok) throw new Error(result.message);
  return result.document;
}

function expectHistory(result: EditorHistoryResult): PptxSourceModel {
  if (!result.ok) throw new Error(result.reason);
  return result.document;
}

function firstShape(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0].shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("shape not found");
  return shape;
}

function firstParagraph(source: PptxSourceModel) {
  return firstShape(source).textBody!.paragraphs[0];
}

function firstTextBody(source: PptxSourceModel) {
  return firstShape(source).textBody!;
}

function firstRun(source: PptxSourceModel) {
  return firstParagraph(source).runs[0];
}

function shapeWithoutTransform(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0].shapes.find(
    (node): node is SourceShape => node.kind === "shape" && node.transform === undefined,
  );
  if (shape === undefined) throw new Error("shape without transform not found");
  return shape;
}

function requireShape(shape: SourceShapeNode | undefined): SourceShapeNode & {
  readonly transform: NonNullable<SourceShapeNode["transform"]>;
} {
  if (shape === undefined || shape.kind === "raw" || shape.transform === undefined) {
    throw new Error("transform shape not found");
  }
  return shape;
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("handle not found");
  return handle;
}
