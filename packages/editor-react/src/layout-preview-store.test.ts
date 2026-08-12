import { asPartPath, asRawSidecarId } from "@pptx-glimpse/document";
import type {
  ConversionDiagnostic,
  PptxEditorTemplatePreviewResult,
  SourceHandle,
} from "pptx-glimpse";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";

import { LayoutThumbnail } from "./EditorLayoutPicker.js";
import { LayoutPreviewStore } from "./layout-preview-store.js";

const first: SourceHandle = { partPath: asPartPath("ppt/slideLayouts/slideLayout1.xml") };
const second: SourceHandle = { partPath: asPartPath("ppt/slideLayouts/slideLayout2.xml") };

describe("LayoutPreviewStore", () => {
  it("caches successful previews and does not schedule the same handle twice", async () => {
    const load = vi.fn((handle: SourceHandle) => Promise.resolve(success(handle)));
    const store = new LayoutPreviewStore(load);

    store.load([first, first]);
    expect(store.get(first)).toEqual({ status: "loading" });
    await settled();
    expect(store.get(first)).toEqual({ status: "ready", svg: "<svg />" });

    store.load([first]);
    expect(load).toHaveBeenCalledTimes(1);
  });

  it("keeps handles with different raw sidecars in distinct cache entries", async () => {
    const load = vi.fn((handle: SourceHandle) => Promise.resolve(success(handle)));
    const store = new LayoutPreviewStore(load);
    const firstSidecar = { ...first, rawSidecarIds: [asRawSidecarId("first")] };
    const secondSidecar = { ...first, rawSidecarIds: [asRawSidecarId("second")] };

    store.load([firstSidecar, secondSidecar]);
    await settled();

    expect(load).toHaveBeenCalledTimes(2);
    expect(store.get(firstSidecar)).toEqual({ status: "ready", svg: "<svg />" });
    expect(store.get(secondSidecar)).toEqual({ status: "ready", svg: "<svg />" });
  });

  it("keeps usable successful SVGs despite unsupported warning diagnostics", async () => {
    const unsupported: ConversionDiagnostic = {
      source: "renderer-adapter",
      severity: "warning",
      code: "pptx-computed-view-adapter.unsupported-test",
      message: "unsupported",
    };
    const store = new LayoutPreviewStore(() =>
      Promise.resolve({ ...success(first), diagnostics: [unsupported] }),
    );

    store.load([first]);
    await settled();

    expect(store.get(first)).toEqual({ status: "ready", svg: "<svg />" });
  });

  it("renders successful SVG previews as non-interactive image data", () => {
    const markup = renderToStaticMarkup(
      LayoutThumbnail({
        name: "Linked layout",
        preview: {
          status: "ready",
          svg: '<svg><a href="https://example.com"><rect /></a></svg>',
        },
      }),
    );

    expect(markup).toContain('<img alt="Linked layout preview"');
    expect(markup).toContain('src="data:image/svg+xml;charset=utf-8,');
    expect(markup).not.toContain('<a href="https://example.com"');
  });

  it("uses deterministic fallback states for lookup, invalid SVG, and runtime failures", async () => {
    const results: Array<PptxEditorTemplatePreviewResult | Error> = [
      { ok: false, code: "preview-handle-not-found", message: "missing", handle: first },
      { ...success(second), svg: "not an svg" },
      { ...success(first), svg: "  " },
      new Error("renderer failed"),
    ];
    const store = new LayoutPreviewStore(() => {
      const result = results.shift();
      if (result instanceof Error) return Promise.reject(result);
      if (result === undefined) return Promise.reject(new Error("missing test result"));
      return Promise.resolve(result);
    }, 1);
    const third: SourceHandle = { partPath: asPartPath("ppt/slideLayouts/slideLayout3.xml") };
    const fourth: SourceHandle = { partPath: asPartPath("ppt/slideLayouts/slideLayout4.xml") };

    store.load([first, second, third, fourth]);
    await settled(8);

    expect(store.get(first)).toEqual({ status: "fallback", message: "Preview unavailable" });
    expect(store.get(second)).toEqual({ status: "fallback", message: "Preview unavailable" });
    expect(store.get(third)).toEqual({ status: "fallback", message: "Preview unavailable" });
    expect(store.get(fourth)).toEqual({ status: "fallback", message: "Preview failed" });

    const fallbackMarkup = renderToStaticMarkup(
      LayoutThumbnail({ name: "Missing layout", preview: store.get(first) }),
    );
    expect(fallbackMarkup).toContain('data-thumbnail-state="fallback"');
    expect(fallbackMarkup).toContain("Preview unavailable");
  });

  it("starts queued work when a concurrency slot becomes available", async () => {
    const resolvers: Array<(result: PptxEditorTemplatePreviewResult) => void> = [];
    const load = vi.fn(
      (_handle: SourceHandle) =>
        new Promise<PptxEditorTemplatePreviewResult>((resolve) => {
          resolvers.push((result) => resolve(result));
        }),
    );
    const store = new LayoutPreviewStore(load, 1);

    store.load([first, second]);
    expect(load).toHaveBeenCalledTimes(1);
    resolvers[0]?.(success(first));
    await settled();
    expect(load).toHaveBeenCalledTimes(2);
    resolvers[1]?.(success(second));
    await settled();
    expect(store.get(second)).toEqual({ status: "ready", svg: "<svg />" });
  });

  it("ignores completion after its session lifecycle is disposed", async () => {
    let resolvePreview: (result: PptxEditorTemplatePreviewResult) => void = () => {};
    const load = vi.fn(
      () =>
        new Promise<PptxEditorTemplatePreviewResult>((resolve) => {
          resolvePreview = resolve;
        }),
    );
    const store = new LayoutPreviewStore(load, 1);

    store.load([first, second]);
    store.dispose();
    resolvePreview(success(first));
    await settled();

    expect(load).toHaveBeenCalledTimes(1);
    expect(store.get(first)).toBeUndefined();
    expect(store.get(second)).toBeUndefined();
  });
});

function success(handle: SourceHandle): PptxEditorTemplatePreviewResult {
  return { ok: true, targetKind: "layout", handle, svg: "<svg />", diagnostics: [] };
}

async function settled(turns = 3): Promise<void> {
  for (let index = 0; index < turns; index += 1) await Promise.resolve();
}
