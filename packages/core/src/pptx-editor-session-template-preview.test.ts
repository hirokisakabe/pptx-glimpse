import { asPartPath } from "@pptx-glimpse/document";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";
import { createPptxEditorSession } from "./index.js";
import { createPptxEditorSessionFactory } from "./pptx-editor-session.js";
import { buildTemplatePreviewFixture } from "./pptx-editor-session.test-helpers.js";

describe("PptxEditorSession - master/layout preview", () => {
  it("renders one catalog target with inherited template content and stable diagnostics", async () => {
    const editor = await createPptxEditorSession(await buildTemplatePreviewFixture(), {
      skipSystemFonts: true,
    });
    const master = editor.layoutCatalog[0];
    const layout = master?.layouts[0];
    if (master === undefined || layout === undefined) throw new Error("catalog target missing");

    const masterPreview = await editor.previewLayoutCatalogTarget(master.handle);
    expect(masterPreview).toMatchObject({ ok: true, targetKind: "master", handle: master.handle });
    if (!masterPreview.ok) throw new Error(masterPreview.message);
    expect(masterPreview.svg).toContain('aria-label="Master normal shape"');
    expect(masterPreview.svg).toContain("MASTER PROMPT");
    expect(masterPreview.svg).not.toContain('aria-label="Layout normal shape"');
    expect(masterPreview.svg).not.toContain("REAL SLIDE USER CONTENT");

    const layoutPreview = await editor.previewLayoutCatalogTarget(layout.handle);
    expect(layoutPreview).toMatchObject({ ok: true, targetKind: "layout", handle: layout.handle });
    if (!layoutPreview.ok) throw new Error(layoutPreview.message);
    expect(layoutPreview.svg).toContain('aria-label="Master normal shape"');
    expect(layoutPreview.svg).toContain('aria-label="Layout normal shape"');
    expect(layoutPreview.svg).toContain("LAYOUT PROMPT");
    expect(layoutPreview.svg).not.toContain("MASTER PROMPT");
    expect(layoutPreview.svg).not.toContain("REAL SLIDE USER CONTENT");
    expect(layoutPreview.svg).toContain("#1155aa");
    expect(layoutPreview.diagnostics).toContainEqual(
      expect.objectContaining({
        source: "renderer-adapter",
        code: "pptx-computed-view-adapter.raw-element-skipped",
        sourcePartPath: layout.handle.partPath,
      }),
    );
    expect(
      layoutPreview.diagnostics.every((diagnostic) => diagnostic.slideNumber === undefined),
    ).toBe(true);
  });

  it("does not change document identity, selection, history, slide cache, or serialization", async () => {
    const editor = await createPptxEditorSession(await buildTemplatePreviewFixture(), {
      skipSystemFonts: true,
    });
    const shape = editor.shapes(1)[0];
    const run = shape?.textBody?.paragraphs[0]?.runs[0];
    const layoutHandle = editor.layoutCatalog[0]?.layouts[0]?.handle;
    if (shape?.handle === undefined || run?.handle === undefined || layoutHandle === undefined) {
      throw new Error("preview non-mutation fixture is incomplete");
    }
    editor.selectShape(shape.handle);
    await editor.apply({
      kind: "replaceTextRunPlainText",
      handle: run.handle,
      text: "EDITED USER",
    });

    const documentBefore = editor.document;
    const selectionBefore = editor.selection;
    const historyBefore = editor.history;
    const slidesBefore = editor.slides;
    const serializedBefore = editor.save().pptx;

    const preview = await editor.previewLayoutCatalogTarget(layoutHandle);

    expect(preview.ok).toBe(true);
    expect(editor.document).toBe(documentBefore);
    expect(editor.selection).toEqual(selectionBefore);
    expect(editor.history).toEqual(historyBefore);
    expect(editor.slides).toBe(slidesBefore);
    expect(await unzipFileContents(editor.save().pptx)).toEqual(
      await unzipFileContents(serializedBefore),
    );
    if (preview.ok) expect(preview.svg).not.toContain("EDITED USER");
  });

  it("returns stable missing and ambiguous catalog-handle failures", async () => {
    const editor = await createPptxEditorSession(await buildTemplatePreviewFixture(), {
      skipSystemFonts: true,
    });
    const missingHandle = { partPath: asPartPath("ppt/slideLayouts/missing.xml") };
    await expect(editor.previewLayoutCatalogTarget(missingHandle)).resolves.toEqual({
      ok: false,
      code: "preview-handle-not-found",
      message: "No slide master or layout matches handle 'ppt/slideLayouts/missing.xml'.",
      handle: missingHandle,
    });

    const master = editor.document.slideMasters[0];
    const layoutPath = master?.layoutPartPaths[0];
    if (master === undefined || layoutPath === undefined) throw new Error("layout target missing");
    Object.defineProperty(editor.document, "slideMasters", {
      value: [{ ...master, layoutPartPaths: [layoutPath, layoutPath] }],
    });
    const ambiguousHandle = { partPath: layoutPath };
    await expect(editor.previewLayoutCatalogTarget(ambiguousHandle)).resolves.toEqual({
      ok: false,
      code: "preview-handle-ambiguous",
      message: `Multiple slide masters or layouts match handle '${layoutPath}'.`,
      handle: ambiguousHandle,
    });
  });

  it("wraps preview renderer failures without changing editor state", async () => {
    const cause = new Error("preview renderer unavailable");
    const createSession = createPptxEditorSessionFactory(
      renderPptxSourceModelToSvg,
      undefined,
      () => Promise.reject(cause),
    );
    const editor = await createSession(await buildTemplatePreviewFixture(), {
      skipSystemFonts: true,
    });
    const handle = editor.layoutCatalog[0]?.layouts[0]?.handle;
    if (handle === undefined) throw new Error("layout missing");
    const documentBefore = editor.document;
    const historyBefore = editor.history;

    await expect(editor.previewLayoutCatalogTarget(handle)).rejects.toMatchObject({
      name: "PptxEditorError",
      code: "render-failed",
      message: "Failed to render master or layout preview: preview renderer unavailable",
      cause,
    });
    expect(editor.document).toBe(documentBefore);
    expect(editor.history).toEqual(historyBefore);
  });

  it.each([
    ["missing", "part-missing"],
    ["unreadable", "part-root-unexpected"],
  ] as const)(
    "keeps %s referenced-theme diagnostics and excludes unrelated slide diagnostics",
    async (themeState, expectedCode) => {
      const editor = await createPptxEditorSession(
        await buildThemeDiagnosticPreviewFixture(themeState),
        { skipSystemFonts: true },
      );
      const layoutHandle = editor.layoutCatalog[0]?.layouts[0]?.handle;
      if (layoutHandle === undefined) throw new Error("layout missing");

      const preview = await editor.previewLayoutCatalogTarget(layoutHandle);

      expect(preview.ok).toBe(true);
      if (!preview.ok) throw new Error(preview.message);
      expect(preview.diagnostics).toContainEqual(
        expect.objectContaining({
          source: "document",
          code: expectedCode,
          sourcePartPath: "ppt/theme/theme1.xml",
        }),
      );
      expect(preview.diagnostics).not.toContainEqual(
        expect.objectContaining({
          source: "document",
          sourcePartPath: "ppt/slides/slide1.xml",
        }),
      );
    },
  );

  it("exposes identical Node and browser values for the same catalog handle", async () => {
    const input = await buildTemplatePreviewFixture();
    const browserEntry = await import("./browser.js");
    const nodeEditor = await createPptxEditorSession(input, { skipSystemFonts: true });
    const browserEditor = await browserEntry.createPptxEditorSession(input, {
      skipSystemFonts: true,
    });
    const nodeHandle = nodeEditor.layoutCatalog[0]?.layouts[0]?.handle;
    const browserHandle = browserEditor.layoutCatalog[0]?.layouts[0]?.handle;
    if (nodeHandle === undefined || browserHandle === undefined) throw new Error("layout missing");

    const nodePreview = await nodeEditor.previewLayoutCatalogTarget(nodeHandle);
    const browserPreview = await browserEditor.previewLayoutCatalogTarget(browserHandle);

    expect(browserPreview).toEqual(nodePreview);
    expect(
      await browserEditor.previewLayoutCatalogTarget({ partPath: asPartPath("missing") }),
    ).toEqual(await nodeEditor.previewLayoutCatalogTarget({ partPath: asPartPath("missing") }));
  });
});

async function buildThemeDiagnosticPreviewFixture(
  themeState: "missing" | "unreadable",
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildTemplatePreviewFixture());
  if (themeState === "missing") {
    zip.remove("ppt/theme/theme1.xml");
  } else {
    zip.file("ppt/theme/theme1.xml", "<unexpected-theme-root/>");
  }
  zip.remove("ppt/slides/slide1.xml");
  return zip.generateAsync({ type: "uint8array" });
}

async function unzipFileContents(pptx: Uint8Array): Promise<Record<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(pptx);
  const files = Object.entries(zip.files)
    .filter(([, file]) => !file.dir)
    .sort(([leftPath], [rightPath]) => leftPath.localeCompare(rightPath));
  return Object.fromEntries(
    await Promise.all(
      files.map(async ([path, file]) => [path, await file.async("uint8array")] as const),
    ),
  );
}
