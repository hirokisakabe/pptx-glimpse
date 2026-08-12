import { asPartPath } from "@pptx-glimpse/document";
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

  it("uses deterministic fallback states for lookup, runtime, and unsupported failures", async () => {
    const unsupported: ConversionDiagnostic = {
      source: "renderer-adapter",
      severity: "warning",
      code: "pptx-computed-view-adapter.unsupported-test",
      message: "unsupported",
    };
    const results: Array<PptxEditorTemplatePreviewResult | Error> = [
      { ok: false, code: "preview-handle-not-found", message: "missing", handle: first },
      { ...success(second), diagnostics: [unsupported] },
      new Error("renderer failed"),
    ];
    const store = new LayoutPreviewStore(() => {
      const result = results.shift();
      if (result instanceof Error) return Promise.reject(result);
      if (result === undefined) return Promise.reject(new Error("missing test result"));
      return Promise.resolve(result);
    }, 1);
    const third: SourceHandle = { partPath: asPartPath("ppt/slideLayouts/slideLayout3.xml") };

    store.load([first, second, third]);
    await settled(6);

    expect(store.get(first)).toEqual({ status: "fallback", message: "Preview unavailable" });
    expect(store.get(second)).toEqual({ status: "fallback", message: "Preview unsupported" });
    expect(store.get(third)).toEqual({ status: "fallback", message: "Preview failed" });

    const fallbackMarkup = renderToStaticMarkup(
      LayoutThumbnail({ name: "Missing layout", preview: store.get(first) }),
    );
    expect(fallbackMarkup).toContain('data-thumbnail-state="fallback"');
    expect(fallbackMarkup).toContain("Preview unavailable");
  });

  it("limits concurrent work and ignores completion after its session lifecycle is disposed", async () => {
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
    store.dispose();
    resolvers[0]?.(success(first));
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
