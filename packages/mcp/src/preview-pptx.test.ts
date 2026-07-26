import type { PngConversionReport } from "pptx-glimpse";
import { describe, expect, it, vi } from "vitest";

import type { PreviewPptxDependencies } from "./preview-pptx.js";
import { previewPptx } from "./preview-pptx.js";

function report(slideNumbers: number[]): PngConversionReport {
  return {
    slides: slideNumbers.map((slideNumber) => ({
      slideNumber,
      png: new Uint8Array([0x89, 0x50, 0x4e, slideNumber]),
      width: 960,
      height: 540,
    })),
    diagnostics: [
      {
        source: "renderer",
        severity: "warning",
        code: "renderer.test-warning",
        message: "test warning",
      },
    ],
    supportCoverage: {
      overall: {
        inputElements: slideNumbers.length,
        outputElements: slideNumbers.length,
        skippedElements: 0,
        unresolvedElements: 0,
        fallbackElements: 0,
        warnings: 1,
      },
      slides: slideNumbers.map((slideNumber) => ({
        slideNumber,
        inputElements: 1,
        outputElements: 1,
        skippedElements: 0,
        unresolvedElements: 0,
        fallbackElements: 0,
        warnings: 0,
      })),
    },
  };
}

function dependencies(
  convert: PreviewPptxDependencies["convertPptxToPng"] = () => Promise.resolve(report([1, 2])),
): PreviewPptxDependencies {
  return {
    readFile: () => Promise.resolve(new Uint8Array([1, 2, 3])),
    convertPptxToPng: convert,
  };
}

function expectTextContent(result: Awaited<ReturnType<typeof previewPptx>>, text: string): void {
  const content = result.content[0];
  expect(content?.type).toBe("text");
  if (content?.type !== "text") {
    throw new Error("Expected text content");
  }
  expect(content.text).toContain(text);
}

describe("previewPptx", () => {
  it("returns every converted slide as PNG image content with structured metadata", async () => {
    const result = await previewPptx({ filePath: "/tmp/deck.pptx" }, dependencies());

    expect(result.isError).not.toBe(true);
    expect(result.content).toEqual([
      { type: "image", data: "iVBOAQ==", mimeType: "image/png" },
      { type: "image", data: "iVBOAg==", mimeType: "image/png" },
    ]);
    expect(result.structuredContent).toMatchObject({
      slides: [
        { slideNumber: 1, contentIndex: 0, width: 960, height: 540 },
        { slideNumber: 2, contentIndex: 1, width: 960, height: 540 },
      ],
      diagnostics: { total: 1, info: 0, warnings: 1, errors: 0 },
      supportCoverage: {
        overall: { inputElements: 2, outputElements: 2, warnings: 1 },
        slides: [
          { slideNumber: 1, inputElements: 1, outputElements: 1 },
          { slideNumber: 2, inputElements: 1, outputElements: 1 },
        ],
      },
    });
  });

  it("passes selected 1-based slide numbers to the public converter", async () => {
    const convert = vi.fn<PreviewPptxDependencies["convertPptxToPng"]>((_input, options) =>
      Promise.resolve(report(options?.slides ?? [])),
    );

    const result = await previewPptx(
      { filePath: "/tmp/deck.pptx", slides: [2] },
      dependencies(convert),
    );

    expect(convert).toHaveBeenCalledWith(expect.any(Uint8Array), {
      slides: [2],
      skipSystemFonts: true,
    });
    expect(result.structuredContent).toMatchObject({
      slides: [{ slideNumber: 2, contentIndex: 0 }],
    });
  });

  it.each([
    { slides: [], message: "at least one" },
    { slides: [0], message: "positive 1-based integers" },
    { slides: [1.5], message: "positive 1-based integers" },
  ])("returns a tool error for invalid slide input $slides", async ({ slides, message }) => {
    const result = await previewPptx({ filePath: "/tmp/deck.pptx", slides }, dependencies());

    expect(result).toMatchObject({ isError: true });
    expectTextContent(result, message);
  });

  it("returns a tool error when the file does not exist", async () => {
    const missing = new Error("missing");
    Object.assign(missing, { code: "ENOENT" });
    const result = await previewPptx(
      { filePath: "/tmp/missing.pptx" },
      {
        ...dependencies(),
        readFile: () => Promise.reject(missing),
      },
    );

    expect(result).toMatchObject({ isError: true });
    expect(result.content[0]).toMatchObject({
      type: "text",
      text: "PPTX file not found: /tmp/missing.pptx",
    });
  });

  it("returns a tool error for an out-of-range slide number", async () => {
    const result = await previewPptx(
      { filePath: "/tmp/deck.pptx", slides: [1, 9] },
      dependencies(() => Promise.resolve(report([1]))),
    );

    expect(result).toMatchObject({ isError: true });
    expect(result.content[0]).toMatchObject({
      type: "text",
      text: "Requested slide number could not be rendered or are out of range: 9",
    });
  });

  it("returns a tool error when the file cannot be read", async () => {
    const result = await previewPptx(
      { filePath: "/tmp/unreadable.pptx" },
      {
        ...dependencies(),
        readFile: () => Promise.reject(new Error("permission denied")),
      },
    );

    expect(result).toMatchObject({ isError: true });
    expectTextContent(result, "Unable to read PPTX file '/tmp/unreadable.pptx': permission denied");
  });

  it("returns a tool error when PPTX conversion fails", async () => {
    const result = await previewPptx(
      { filePath: "/tmp/invalid.pptx" },
      dependencies(() => Promise.reject(new Error("Invalid PPTX archive"))),
    );

    expect(result).toMatchObject({ isError: true });
    expectTextContent(result, "Invalid PPTX archive");
  });
});
