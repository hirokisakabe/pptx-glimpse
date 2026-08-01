import {
  asEmu,
  asPartPath,
  asRawSidecarId,
  asSourceNodeId,
  findShapeNodeBySourceHandle,
  readPptx,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";
import {
  BLUE_PNG,
  buildImageReplacementFixture,
  buildTextEditFixture,
  buildTwoSlideFixture,
  buildUnreferencedLayoutFixture,
  expectApplied,
  expectHistory,
  firstImage,
  firstParagraph,
  firstRun,
  firstShape,
  requireHandle,
} from "./index.test-helpers.js";

describe("EditorSession source-node convenience methods", () => {
  it("edits text runs and paragraphs without exposing handles or command objects", async () => {
    const source = readPptx(await buildTextEditFixture());
    const run = firstRun(source);
    const paragraph = firstParagraph(source);
    const runSession = createEditorSession(source);

    expectApplied(runSession.replaceTextRunPlainText(run, "First edit"));
    expectApplied(runSession.setTextRunProperties(run, { bold: true }));
    expectApplied(runSession.clearTextRunProperties(run, ["bold"]));
    expectApplied(runSession.replaceTextRunPlainText(run, "Edited through stale node"));

    expect(firstRun(runSession.document).text).toBe("Edited through stale node");
    expect(runSession.undoDepth).toBe(3);
    expect(firstRun(expectHistory(runSession.undo())).text).toBe("First edit");
    expect(firstRun(expectHistory(runSession.redo())).text).toBe("Edited through stale node");

    const paragraphSession = createEditorSession(source);
    expectApplied(paragraphSession.setParagraphProperties(paragraph, { align: "center" }));
    expectApplied(paragraphSession.clearParagraphProperties(paragraph, ["align"]));
    expectApplied(paragraphSession.replaceParagraphPlainText(paragraph, "Paragraph edit"));

    expect(firstParagraph(paragraphSession.document).runs).toHaveLength(1);
    expect(firstRun(paragraphSession.document).text).toBe("Paragraph edit");
    expect(paragraphSession.undoDepth).toBe(2);
  });

  it("edits and deletes a shape through a stale source node while preserving selection history", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const shape = firstShape(source);
    const shapeHandle = requireHandle(shape.handle);

    expect(session.selectShape(shapeHandle)).toMatchObject({ ok: true });
    expectApplied(session.moveShape(shape, asEmu(1000), asEmu(2000)));
    expectApplied(session.resizeShape(shape, asEmu(3000), asEmu(4000)));
    expectApplied(
      session.setShapeTransform(shape, {
        offsetX: asEmu(5000),
        offsetY: asEmu(6000),
        width: asEmu(7000),
        height: asEmu(8000),
      }),
    );
    expectApplied(
      session.setShapeFill(shape, { kind: "solid", color: { kind: "srgb", hex: "00AA44" } }),
    );
    expectApplied(session.setShapeOutline(shape, { fill: { kind: "none" } }));
    expectApplied(session.deleteShape(shape));

    expect(findShapeNodeBySourceHandle(session.document, shapeHandle)).toBeUndefined();
    expect(session.selection).toBeUndefined();
    expect(session.undoDepth).toBe(6);
    expect(findShapeNodeBySourceHandle(expectHistory(session.undo()), shapeHandle)).toBeDefined();
  });

  it("targets slides for additions and topology edits and images for replacement", async () => {
    const textSource = readPptx(await buildTextEditFixture());
    const addSession = createEditorSession(textSource);
    const slide = textSource.slides[0];

    expectApplied(
      addSession.addTextBox(slide, {
        offsetX: asEmu(100),
        offsetY: asEmu(200),
        width: asEmu(300),
        height: asEmu(400),
        text: "Convenient",
      }),
    );
    expectApplied(
      addSession.addConnector(slide, {
        preset: "straightConnector1",
        offsetX: asEmu(500),
        offsetY: asEmu(600),
        width: asEmu(700),
        height: asEmu(800),
      }),
    );
    expect(addSession.undoDepth).toBe(2);

    const imageSource = readPptx(await buildImageReplacementFixture());
    const imageSession = createEditorSession(imageSource);
    const imageResult = imageSession.replaceImage(firstImage(imageSource), BLUE_PNG);
    expectApplied(imageResult);
    expect(imageResult.ok && imageResult.warnings).toBeUndefined();

    const slideSource = readPptx(await buildTwoSlideFixture());
    const slideSession = createEditorSession(slideSource);
    const first = slideSource.slides[0];
    const slideWithNonIdentityHandleMetadata = {
      ...first,
      handle: {
        ...requireHandle(first.handle),
        rawSidecarIds: [asRawSidecarId("non-identity-metadata")],
      },
    };
    expectApplied(slideSession.moveSlide(slideWithNonIdentityHandleMetadata, { toIndex: 1 }));
    expectApplied(slideSession.duplicateSlide(first));
    expectApplied(slideSession.moveSlide(first, { toIndex: 2 }));
    expectApplied(slideSession.deleteSlide(first));
    expect(slideSession.document.slides).toHaveLength(2);
    expect(slideSession.undoDepth).toBe(4);

    const layoutSource = readPptx(await buildUnreferencedLayoutFixture());
    const layoutSession = createEditorSession(layoutSource);
    expectApplied(
      layoutSession.addEmptySlideFromLayout({
        layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout2.xml"),
      }),
    );
    expect(layoutSession.document.slides).toHaveLength(2);
  });

  it("rejects missing, wrong-type, and absent source nodes atomically", async () => {
    const source = readPptx(await buildTextEditFixture());
    const session = createEditorSession(source);
    const run = firstRun(source);
    const shape = firstShape(source);
    expectApplied(session.replaceTextRunPlainText(run, "History setup"));
    expectHistory(session.undo());
    const before = session.document;
    const missingHandleRun = { ...run, handle: undefined };
    const absentRun = {
      ...run,
      handle: {
        partPath: asPartPath("ppt/slides/missing.xml"),
        nodeId: asSourceNodeId("missing-run"),
      },
    };

    const missingHandleResult = session.replaceTextRunPlainText(missingHandleRun, "Rejected");
    expect(missingHandleResult).toMatchObject({ ok: false, code: "invalid-command" });
    expect(!missingHandleResult.ok && missingHandleResult.message).toContain(
      "does not have a handle",
    );
    expect(!missingHandleResult.ok && missingHandleResult.cause).toBeUndefined();
    const absentResult = session.replaceTextRunPlainText(absentRun, "Rejected");
    expect(absentResult).toMatchObject({ ok: false, code: "invalid-command" });
    expect(!absentResult.ok && absentResult.message).toContain("current EditorSession document");
    // @ts-expect-error Exercise runtime rejection when untyped JavaScript passes a text run.
    const wrongParagraphResult = session.replaceParagraphPlainText(run, "Rejected");
    expect(wrongParagraphResult).toMatchObject({ ok: false, code: "invalid-command" });
    expect(!wrongParagraphResult.ok && wrongParagraphResult.message).toContain("wrong target type");
    // @ts-expect-error Exercise runtime rejection when untyped JavaScript passes a shape.
    const wrongImageResult = session.replaceImage(shape, BLUE_PNG);
    expect(wrongImageResult).toMatchObject({ ok: false, code: "invalid-command" });
    expect(!wrongImageResult.ok && wrongImageResult.message).toContain("wrong target type");

    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(1);
  });
});
