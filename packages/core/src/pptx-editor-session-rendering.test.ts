import { asEmu, readPptx } from "@pptx-glimpse/document";
import { describe, expect, it, vi } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";
import { createPptxEditorSession, PptxEditorSession } from "./index.js";
import { createPptxEditorSessionFactory } from "./pptx-editor-session.js";
import {
  buildShapeFixture,
  buildTwoSlideFixture,
  firstShape,
  firstText,
} from "./pptx-editor-session.test-helpers.js";

const nodeFontMocks = vi.hoisted(() => ({
  createOpentypeSetupFromSystem: vi.fn().mockResolvedValue(null),
}));

vi.mock("@pptx-glimpse/renderer/node", () => ({
  createOpentypeSetupFromSystem: nodeFontMocks.createOpentypeSetupFromSystem,
}));

describe("PptxEditorSession - rendering", () => {
  it("edits, rerenders, and saves Uint8Array PPTX bytes in Node.js", async () => {
    const editor = await createPptxEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const shape = editor.shapes(1)[0];
    if (shape?.handle === undefined) throw new Error("shape handle not found");

    expect(editor.slides).toHaveLength(1);
    expect(shape.bounds).toEqual({ x: 96, y: 192, width: 288, height: 96 });
    expect(shape.textBody).toMatchObject({
      paragraphs: [
        {
          runs: [
            {
              text: "Original",
              properties: {
                fontSize: 24,
                typeface: "Aptos",
              },
            },
          ],
        },
      ],
    });
    expect(shape.textBody?.paragraphs[0]?.handle).toBeDefined();
    expect(shape.textBody?.paragraphs[0]?.runs[0]?.handle).toBeDefined();
    const runHandle = shape.textBody?.paragraphs[0]?.runs[0]?.handle;
    if (runHandle === undefined) throw new Error("text run handle not found");

    await editor.apply({
      kind: "setShapeTransform",
      handle: shape.handle,
      offsetX: asEmu(120 * 9525),
      offsetY: asEmu(208 * 9525),
      width: asEmu(336 * 9525),
      height: asEmu(120 * 9525),
    });
    await editor.apply({
      kind: "replaceTextRunPlainText",
      handle: runHandle,
      text: "Node edited",
    });

    expect(editor.history.undoDepth).toBe(2);
    expect(editor.slides[0]?.svg).toContain("Node edited");

    expect((await editor.undo()).history).toMatchObject({ canRedo: true, undoDepth: 1 });
    expect(firstText(editor.document)).toBe("Original");
    expect((await editor.redo()).history).toMatchObject({ canUndo: true, redoDepth: 0 });
    expect(firstText(editor.document)).toBe("Node edited");

    const saved = readPptx(editor.save().pptx);
    expect(firstText(saved)).toBe("Node edited");
    expect(firstShape(saved).transform).toMatchObject({
      offsetX: 120 * 9525,
      offsetY: 208 * 9525,
      width: 336 * 9525,
      height: 120 * 9525,
    });
  });

  it("uses the Node font loader selected by the main entry", async () => {
    nodeFontMocks.createOpentypeSetupFromSystem.mockClear();

    await createPptxEditorSession(await buildShapeFixture(), {
      fontDirs: ["/app/fonts"],
      skipSystemFonts: true,
    });

    expect(nodeFontMocks.createOpentypeSetupFromSystem).toHaveBeenCalledWith(
      ["/app/fonts"],
      undefined,
      true,
    );
  });

  it("keeps Node and browser static session factories entry-specific", async () => {
    const browserEntry = await import("./browser.js");
    nodeFontMocks.createOpentypeSetupFromSystem.mockClear();
    const input = await buildShapeFixture();
    const renderOptions = {
      fontDirs: ["/static-node-fonts"],
      skipSystemFonts: true,
    };

    const browserEditor = await browserEntry.PptxEditorSession.create(input, renderOptions);
    expect(browserEditor).toBeInstanceOf(browserEntry.PptxEditorSession);
    expect(nodeFontMocks.createOpentypeSetupFromSystem).not.toHaveBeenCalled();

    const nodeEditor = await PptxEditorSession.create(input, renderOptions);
    expect(nodeEditor).toBeInstanceOf(PptxEditorSession);
    expect(nodeFontMocks.createOpentypeSetupFromSystem).toHaveBeenCalledOnce();
  });

  it("keeps Node and browser editor factories isolated regardless of import order", async () => {
    const browserEntry = await import("./browser.js");
    nodeFontMocks.createOpentypeSetupFromSystem.mockClear();
    const input = await buildShapeFixture();
    const renderOptions = {
      fontDirs: ["/isolated-node-fonts"],
      skipSystemFonts: true,
    };

    await browserEntry.createPptxEditorSession(input, renderOptions);
    expect(nodeFontMocks.createOpentypeSetupFromSystem).not.toHaveBeenCalled();

    await createPptxEditorSession(input, renderOptions);
    expect(nodeFontMocks.createOpentypeSetupFromSystem).toHaveBeenCalledOnce();
  });

  it("applies command batches as one history entry and renders once", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    const createTestEditorSession = createPptxEditorSessionFactory((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    const editor = await createTestEditorSession(await buildShapeFixture(), {
      skipSystemFonts: true,
    });
    const run = editor.shapes(1)[0]?.textBody?.paragraphs[0]?.runs[0];
    if (run?.handle === undefined) throw new Error("text run handle not found");

    const response = await editor.applyAll([
      {
        kind: "replaceTextRunPlainText",
        handle: run.handle,
        text: "Batch edited",
      },
      {
        kind: "setTextRunProperties",
        handle: run.handle,
        properties: { bold: true },
      },
    ]);

    expect(renderCalls).toEqual([undefined, [1]]);
    expect(response.history).toMatchObject({ undoDepth: 1, canUndo: true });
    expect(firstText(editor.document)).toBe("Batch edited");
    expect(firstShape(editor.document).textBody?.paragraphs[0]?.runs[0]?.properties?.bold).toBe(
      true,
    );

    await editor.undo();
    expect(firstText(editor.document)).toBe("Original");
    expect(
      firstShape(editor.document).textBody?.paragraphs[0]?.runs[0]?.properties?.bold,
    ).toBeUndefined();
  });

  it("rerenders only the affected slides for applyAll, undo, and redo", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    let renderGeneration = 0;
    const createTestEditorSession = createPptxEditorSessionFactory(async (source, options) => {
      renderGeneration += 1;
      renderCalls.push(options?.slides);
      const report = await renderPptxSourceModelToSvg(source, options);
      return {
        ...report,
        slides: report.slides.map((slide) => ({
          ...slide,
          svg: `${slide.svg}<!-- render:${String(renderGeneration)} -->`,
        })),
      };
    });
    const editor = await createTestEditorSession(await buildTwoSlideFixture(), {
      skipSystemFonts: true,
    });
    const initialFirstSvg = editor.slides[0]?.svg;
    const firstRun = editor.shapes(1)[0]?.textBody?.paragraphs[0]?.runs[0];
    const secondRun = editor.shapes(2)[0]?.textBody?.paragraphs[0]?.runs[0];
    if (firstRun?.handle === undefined || secondRun?.handle === undefined) {
      throw new Error("text run handles not found");
    }

    await editor.apply({
      kind: "replaceTextRunPlainText",
      handle: secondRun.handle,
      text: "Second edited",
    });
    expect(renderCalls).toEqual([undefined, [2]]);
    expect(editor.slides[0]?.svg).toBe(initialFirstSvg);
    expect(editor.slides[1]?.svg).toContain("Second edited");

    await editor.applyAll([
      {
        kind: "replaceTextRunPlainText",
        handle: firstRun.handle,
        text: "First batch edited",
      },
      {
        kind: "replaceTextRunPlainText",
        handle: secondRun.handle,
        text: "Second batch edited",
      },
      {
        kind: "setTextRunProperties",
        handle: secondRun.handle,
        properties: { bold: true },
      },
    ]);
    expect(renderCalls.at(-1)).toEqual([1, 2]);
    expect(editor.slides[0]?.svg).toContain("First batch edited");
    expect(editor.slides[1]?.svg).toContain("Second batch edited");

    await editor.undo();
    expect(renderCalls.at(-1)).toEqual([1, 2]);
    expect(editor.slides[0]?.svg).toContain("First");
    expect(editor.slides[1]?.svg).toContain("Second edited");

    await editor.redo();
    expect(renderCalls.at(-1)).toEqual([1, 2]);
    expect(editor.slides[0]?.svg).toContain("First batch edited");
    expect(editor.slides[1]?.svg).toContain("Second batch edited");
  });
});
