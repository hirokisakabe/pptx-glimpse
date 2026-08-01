import { asPartPath, readPptx, writePptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";
import {
  buildTwoSlideFixture,
  buildUnreferencedLayoutFixture,
  expectApplied,
  expectHistory,
  requireHandle,
} from "./index.test-helpers.js";

describe("EditorSession slide topology commands", () => {
  it("adds an empty slide from a layout as one undoable command and persists it", async () => {
    const source = readPptx(await buildUnreferencedLayoutFixture());
    const session = createEditorSession(source);
    const added = expectApplied(
      session.apply({
        kind: "addEmptySlideFromLayout",
        layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout2.xml"),
      }),
    );

    expect(added.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(readPptx(writePptx(added)).slides[1]?.layoutPartPath).toBe(
      "ppt/slideLayouts/slideLayout2.xml",
    );
    expect(session.undoDepth).toBe(1);

    const undone = expectHistory(session.undo());
    expect(undone.presentation.slidePartPaths).toEqual(source.presentation.slidePartPaths);
    const redone = expectHistory(session.redo());
    expect(redone.presentation.slidePartPaths).toEqual(added.presentation.slidePartPaths);
  });

  it("duplicates a slide as one undoable command and persists it", async () => {
    const source = readPptx(await buildTwoSlideFixture());
    const session = createEditorSession(source);
    const duplicated = expectApplied(
      session.apply({ kind: "duplicateSlide", handle: requireHandle(source.slides[0].handle) }),
    );

    expect(duplicated.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(readPptx(writePptx(duplicated)).presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(session.undoDepth).toBe(1);

    const undone = expectHistory(session.undo());
    expect(undone.presentation.slidePartPaths).toEqual(source.presentation.slidePartPaths);
    const redone = expectHistory(session.redo());
    expect(redone.presentation.slidePartPaths).toEqual(duplicated.presentation.slidePartPaths);
  });

  it("moves a slide as one undoable command and persists it", async () => {
    const source = readPptx(await buildTwoSlideFixture());
    const session = createEditorSession(source);
    const moved = expectApplied(
      session.apply({
        kind: "moveSlide",
        handle: requireHandle(source.slides[0].handle),
        toIndex: 1,
      }),
    );

    expect(moved.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(readPptx(writePptx(moved)).presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(session.undoDepth).toBe(1);

    const undone = expectHistory(session.undo());
    expect(undone.presentation.slidePartPaths).toEqual(source.presentation.slidePartPaths);
    const redone = expectHistory(session.redo());
    expect(redone.presentation.slidePartPaths).toEqual(moved.presentation.slidePartPaths);
  });

  it("deletes a slide as one undoable command and rejects invalid slide deletes", async () => {
    const source = readPptx(await buildTwoSlideFixture());
    const session = createEditorSession(source);
    const deleted = expectApplied(
      session.apply({ kind: "deleteSlide", handle: requireHandle(source.slides[0].handle) }),
    );

    expect(deleted.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
    expect(readPptx(writePptx(deleted)).presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
    ]);

    const undone = expectHistory(session.undo());
    expect(undone.presentation.slidePartPaths).toEqual(source.presentation.slidePartPaths);
    expectHistory(session.redo());

    const lastSlideReject = session.apply({
      kind: "deleteSlide",
      handle: requireHandle(session.document.slides[0].handle),
    });
    expect(lastSlideReject).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.undoDepth).toBe(1);

    const missingHandleReject = createEditorSession(source).apply({
      kind: "duplicateSlide",
      handle: { partPath: asPartPath("ppt/slides/missing.xml") },
    });
    expect(missingHandleReject).toMatchObject({ ok: false, code: "invalid-command" });
  });
});
