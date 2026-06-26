import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it, vi } from "vitest";

import * as documentExperimental from "../../pptx-glimpse-document/src/experimental.js";
import * as adapterModule from "./cleandoc-renderer-adapter.js";
import { convertPptxToPng, convertPptxToSvg } from "./converter.js";
import {
  convertPptxToPngViaDocumentPath,
  convertPptxToSvgViaDocumentPath,
} from "./experimental-document-renderer.js";

const SELECTED_SHARED_FIXTURES = ["real-basic-theme.pptx", "real-product-page.pptx"] as const;

const DOCUMENT_RENDER_TEST_SCOPE = [
  "readPptx source model",
  "createComputedView cascade projection",
  "CleanDoc computed-view to current renderer model adapter",
  "existing SVG renderer",
  "existing SVG to PNG conversion",
] as const;

const DOCUMENT_RENDER_UNSUPPORTED_SUBSET = [
  "raw shape-tree nodes such as groups and unsupported graphicFrame content",
  "raw background and fill variants outside the CleanDoc adapter subset",
  "unresolved images without a package media payload",
  "elements missing computed transforms",
] as const;

describe("experimental document render path", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("documents the intentionally focused dogfood scope", () => {
    expect(DOCUMENT_RENDER_TEST_SCOPE).toMatchInlineSnapshot(`
      [
        "readPptx source model",
        "createComputedView cascade projection",
        "CleanDoc computed-view to current renderer model adapter",
        "existing SVG renderer",
        "existing SVG to PNG conversion",
      ]
    `);
    expect(DOCUMENT_RENDER_UNSUPPORTED_SUBSET).toMatchInlineSnapshot(`
      [
        "raw shape-tree nodes such as groups and unsupported graphicFrame content",
        "raw background and fill variants outside the CleanDoc adapter subset",
        "unresolved images without a package media payload",
        "elements missing computed transforms",
      ]
    `);
  });

  it.each(SELECTED_SHARED_FIXTURES)(
    "renders selected slides from %s to SVG through the document path",
    async (fixtureName) => {
      const result = await convertPptxToSvgViaDocumentPath(readFixture(fixtureName), {
        slides: [1],
        textOutput: "text",
        skipSystemFonts: true,
      });

      expect(result.slides.map((slide) => slide.slideNumber)).toEqual([1]);
      expect(result.slides[0]?.svg).toContain("<svg");
      expect(result.slides[0]?.svg).toMatch(/viewBox="0 0 \d+ \d+"/);
      expect(result.slides[0]?.svg).toContain("</svg>");
      expectUnsupportedDiagnosticsToStayInScope(result.diagnostics);
    },
  );

  it("does not emit CJK font-context diagnostics when document text fonts are exposed", async () => {
    const result = await convertPptxToSvgViaDocumentPath(readFixture("real-basic-theme.pptx"), {
      slides: [1],
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(result.diagnostics).toEqual([]);
    expectUnsupportedDiagnosticsToStayInScope(result.diagnostics);
  });

  it("connects the document SVG path to the existing PNG conversion", async () => {
    const result = await convertPptxToPngViaDocumentPath(readFixture("real-basic-theme.pptx"), {
      slides: [1],
      width: 240,
      skipSystemFonts: true,
    });

    expect(result.slides.map((slide) => slide.slideNumber)).toEqual([1]);
    const pngSlide = result.slides[0];
    expect(pngSlide).toMatchObject({ width: 240 });
    if (pngSlide === undefined) {
      throw new Error("Expected one PNG slide from the document render path");
    }
    expect([...pngSlide.png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);
    expect(result.diagnostics).toEqual([]);
    expectUnsupportedDiagnosticsToStayInScope(result.diagnostics);
  });

  it("keeps the public SVG converter on its existing default path", async () => {
    const readPptxSpy = vi.spyOn(documentExperimental, "readPptx");
    const adapterSpy = vi.spyOn(adapterModule, "adaptComputedViewToRendererModel");
    const input = readFixture("real-basic-theme.pptx");
    const publicDefault = await convertPptxToSvg(input, {
      slides: [1],
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(publicDefault.map((slide) => slide.slideNumber)).toEqual([1]);
    expect(publicDefault[0]?.svg).toContain("<svg");
    expect(readPptxSpy).not.toHaveBeenCalled();
    expect(adapterSpy).not.toHaveBeenCalled();
  });

  it("keeps the public PNG converter on its existing default path", async () => {
    const readPptxSpy = vi.spyOn(documentExperimental, "readPptx");
    const adapterSpy = vi.spyOn(adapterModule, "adaptComputedViewToRendererModel");
    const input = readFixture("real-basic-theme.pptx");
    const publicDefault = await convertPptxToPng(input, {
      slides: [1],
      width: 240,
      skipSystemFonts: true,
    });

    expect(publicDefault.map((slide) => slide.slideNumber)).toEqual([1]);
    expect(publicDefault[0]).toMatchObject({ width: 240 });
    expect(readPptxSpy).not.toHaveBeenCalled();
    expect(adapterSpy).not.toHaveBeenCalled();
  });
});

function readFixture(name: (typeof SELECTED_SHARED_FIXTURES)[number]): Buffer {
  return readFileSync(fileURLToPath(new URL(`../../../shared-fixtures/${name}`, import.meta.url)));
}

function uniqueDiagnosticCodes(
  diagnostics: readonly { readonly code: string }[],
): readonly string[] {
  return [...new Set(diagnostics.map((diagnostic) => diagnostic.code))].sort();
}

function expectUnsupportedDiagnosticsToStayInScope(
  diagnostics: readonly { readonly code: string }[],
): void {
  expect(uniqueDiagnosticCodes(diagnostics).every(isExpectedDocumentRenderDiagnosticCode)).toBe(
    true,
  );
}

function isExpectedDocumentRenderDiagnosticCode(code: string): boolean {
  return code === "cleandoc-adapter.raw-element-skipped";
}
