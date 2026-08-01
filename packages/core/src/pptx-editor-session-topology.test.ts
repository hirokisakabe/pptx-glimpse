import {
  addEmptySlideFromLayout,
  addShape,
  addSlideLayout,
  asEmu,
  createPptx,
  readPptx,
  setBackground,
  setShapeFill,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";
import { affectedSlidePartPaths, createPptxEditorSessionFactory } from "./pptx-editor-session.js";
import {
  buildLayoutCatalogFixture,
  buildTwoSlideFixture,
  RED_PNG,
} from "./pptx-editor-session.test-helpers.js";

describe("PptxEditorSession - topology", () => {
  it("duplicates and deletes slides with render state and history updates", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    const createTestEditorSession = createPptxEditorSessionFactory((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    const editor = await createTestEditorSession(await buildTwoSlideFixture(), {
      skipSystemFonts: true,
    });
    const firstSlide = editor.slides[0];
    if (firstSlide?.handle === undefined) throw new Error("first slide handle not found");

    const duplicated = await editor.apply({ kind: "duplicateSlide", handle: firstSlide.handle });
    expect(renderCalls).toEqual([undefined, [2]]);
    expect(duplicated.slides).toHaveLength(3);
    expect(duplicated.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
      [1, "ppt/slides/slide1.xml"],
      [2, "ppt/slides/slide3.xml"],
      [3, "ppt/slides/slide2.xml"],
    ]);
    expect(duplicated.slides[0]?.svg).toContain("First");
    expect(duplicated.slides[1]?.svg).toContain("First");
    expect(duplicated.history).toMatchObject({ canUndo: true, undoDepth: 1 });

    const moved = await editor.apply({
      kind: "moveSlide",
      handle: firstSlide.handle,
      toIndex: 2,
    });
    expect(renderCalls).toEqual([undefined, [2]]);
    expect(moved.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
      [1, "ppt/slides/slide3.xml"],
      [2, "ppt/slides/slide2.xml"],
      [3, "ppt/slides/slide1.xml"],
    ]);

    const duplicateSlide = moved.slides[0];
    if (duplicateSlide?.handle === undefined) throw new Error("duplicate slide handle not found");
    const deleted = await editor.apply({ kind: "deleteSlide", handle: duplicateSlide.handle });
    expect(renderCalls).toEqual([undefined, [2]]);
    expect(deleted.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
      [1, "ppt/slides/slide2.xml"],
      [2, "ppt/slides/slide1.xml"],
    ]);
    expect(deleted.history.undoDepth).toBe(3);

    expect((await editor.undo()).slides.map((slide) => slide.handle?.partPath)).toEqual([
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(renderCalls.at(-1)).toEqual([1]);
    const renderCallCountAfterUndo = renderCalls.length;
    expect((await editor.redo()).slides.map((slide) => slide.handle?.partPath)).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(renderCalls).toHaveLength(renderCallCountAfterUndo);

    const added = await editor.apply({
      kind: "addEmptySlideFromLayout",
      layoutPartPath: "ppt/slideLayouts/slideLayout1.xml",
    });
    expect(renderCalls.at(-1)).toEqual([3]);
    expect(added.slides.map((slide) => [slide.slideNumber, slide.handle?.partPath])).toEqual([
      [1, "ppt/slides/slide2.xml"],
      [2, "ppt/slides/slide1.xml"],
      [3, "ppt/slides/slide3.xml"],
    ]);
  });

  it("falls back to rendering all slides when a command change cannot be scoped safely", async () => {
    const before = readPptx(await buildTwoSlideFixture());
    const after = {
      ...before,
      diagnostics: [
        ...before.diagnostics,
        { severity: "warning" as const, code: "unknown-change", message: "unknown change" },
      ],
    };

    expect(affectedSlidePartPaths(before, after)).toBeUndefined();

    const renderCalls: Array<readonly number[] | undefined> = [];
    const createTestEditorSession = createPptxEditorSessionFactory(
      (source, options) => {
        renderCalls.push(options?.slides);
        return renderPptxSourceModelToSvg(source, options);
      },
      () => undefined,
    );
    const editor = await createTestEditorSession(await buildTwoSlideFixture(), {
      skipSystemFonts: true,
    });
    const run = editor.shapes(2)[0]?.textBody?.paragraphs[0]?.runs[0];
    if (run?.handle === undefined) throw new Error("text run handle not found");

    await editor.apply({
      kind: "replaceTextRunPlainText",
      handle: run.handle,
      text: "Fallback edited",
    });

    expect(renderCalls).toEqual([undefined, undefined]);
    expect(editor.slides[0]?.svg).toContain("First");
    expect(editor.slides[1]?.svg).toContain("Fallback edited");
  });

  it("scopes and renders inherited master and layout background changes", async () => {
    const created = createPptx();
    const firstSlide = created.slides[0];
    const firstLayout = created.slideLayouts[0];
    const master = created.slideMasters[0];
    if (
      firstSlide === undefined ||
      firstLayout?.handle === undefined ||
      master?.handle === undefined
    ) {
      throw new Error("createPptx should create a slide, layout, and master");
    }
    let before = addSlideLayout(created, master.handle, { name: "Second layout" });
    const secondLayout = before.slideLayouts.at(-1);
    if (secondLayout === undefined) throw new Error("second layout was not authored");
    before = addEmptySlideFromLayout(before, { layoutPartPath: secondLayout.partPath });
    const secondSlide = before.slides.at(-1);
    if (secondSlide?.handle === undefined) throw new Error("second slide was not authored");
    before = setBackground(before, secondSlide.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "FFFFFF" },
    });

    const afterMaster = setBackground(before, master.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    expect(affectedSlidePartPaths(before, afterMaster)).toEqual(new Set([firstSlide.partPath]));
    const afterMasterImage = setBackground(before, master.handle, {
      kind: "image",
      bytes: RED_PNG,
    });
    expect(affectedSlidePartPaths(before, afterMasterImage)).toEqual(
      new Set([firstSlide.partPath]),
    );
    const renderedMaster = await renderPptxSourceModelToSvg(afterMaster, {
      skipSystemFonts: true,
    });
    expect(renderedMaster.slides[0]?.svg).toContain('fill="#112233"');
    expect(renderedMaster.slides[1]?.svg).toContain('fill="#ffffff"');

    const afterLayout = setBackground(before, firstLayout.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "445566" },
    });
    expect(affectedSlidePartPaths(before, afterLayout)).toEqual(new Set([firstSlide.partPath]));
    const rendered = await renderPptxSourceModelToSvg(afterLayout, { skipSystemFonts: true });
    expect(rendered.slides[0]?.svg).toContain('fill="#445566"');
    expect(rendered.slides[1]?.svg).toContain('fill="#ffffff"');
  });

  it("resolves affected slides for layout and master shape property edits", async () => {
    const source = readPptx(await buildLayoutCatalogFixture());
    const layout = source.slideLayouts.find(
      (candidate) => candidate.partPath === "ppt/slideLayouts/slideLayout2.xml",
    );
    const master = source.slideMasters.find(
      (candidate) => candidate.partPath === "ppt/slideMasters/slideMaster1.xml",
    );
    if (layout?.handle === undefined || master?.handle === undefined) {
      throw new Error("layout or master handle not found");
    }
    const withLayoutShape = addShape(source, layout.handle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
    });
    const layoutShape = withLayoutShape.slideLayouts
      .find((candidate) => candidate.partPath === layout.partPath)
      ?.shapes.at(-1);
    if (layoutShape?.handle === undefined) throw new Error("layout shape handle not found");
    const editedLayout = setShapeFill(withLayoutShape, layoutShape.handle, { kind: "none" });

    expect([...affectedSlidePartPaths(withLayoutShape, editedLayout)!]).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
    ]);

    const withMasterShape = addShape(source, master.handle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
    });
    const masterShape = withMasterShape.slideMasters
      .find((candidate) => candidate.partPath === master.partPath)
      ?.shapes.at(-1);
    if (masterShape?.handle === undefined) throw new Error("master shape handle not found");
    const editedMaster = setShapeFill(withMasterShape, masterShape.handle, { kind: "none" });

    expect([...affectedSlidePartPaths(withMasterShape, editedMaster)!]).toEqual([
      "ppt/slides/slide1.xml",
    ]);
  });

  it("resolves all affected slides when applyAll changes inherited and slide-local content", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    const createTestEditorSession = createPptxEditorSessionFactory((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    const editor = await createTestEditorSession(await buildTwoSlideFixture(), {
      skipSystemFonts: true,
    });
    const layoutHandle = editor.document.slideLayouts[0]?.handle;
    const secondRun = editor.shapes(2)[0]?.textBody?.paragraphs[0]?.runs[0];
    if (layoutHandle === undefined || secondRun?.handle === undefined) {
      throw new Error("layout or text run handle not found");
    }

    await editor.applyAll([
      {
        kind: "addTextBox",
        slideHandle: layoutHandle,
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(914400),
        height: asEmu(914400),
        text: "Inherited edit",
      },
      {
        kind: "replaceTextRunPlainText",
        handle: secondRun.handle,
        text: "Second local edit",
      },
    ]);

    expect(renderCalls).toEqual([undefined, [1, 2]]);
    expect(editor.slides[0]?.svg).toContain(">Inher</");
    expect(editor.slides[1]?.svg).toContain("Second local edit");

    await editor.undo();
    expect(renderCalls).toEqual([undefined, [1, 2], [1, 2]]);
    expect(editor.slides[0]?.svg).not.toContain(">Inher</");
    expect(editor.slides[1]?.svg).toContain("Second");
  });
});
