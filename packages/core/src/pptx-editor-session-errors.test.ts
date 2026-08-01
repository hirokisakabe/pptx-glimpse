import { asSourceNodeId } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";
import { createPptxEditorSession, isPptxEditorError } from "./index.js";
import { configurePptxEditorSessionRenderer } from "./pptx-editor-session.js";
import {
  buildLayoutCatalogFixture,
  buildShapeFixture,
  capturePptxEditorError,
  captureSynchronousPptxEditorError,
  expectErrorCodeAndCause,
} from "./pptx-editor-session.test-helpers.js";

describe("PptxEditorSession - errors", () => {
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
});
