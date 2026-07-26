import { readFile } from "node:fs/promises";

import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { convertPptxToPng, type PngConversionReport, type SupportCoverage } from "pptx-glimpse";

export type PreviewPptxInput = {
  filePath: string;
  slides?: number[];
};

export type PreviewPptxStructuredContent = {
  slides: {
    slideNumber: number;
    contentIndex: number;
    width: number;
    height: number;
  }[];
  diagnostics: {
    total: number;
    info: number;
    warnings: number;
    errors: number;
  };
  supportCoverage: SupportCoverage;
};

export type PreviewPptxDependencies = {
  readFile: (filePath: string) => Promise<Uint8Array>;
  convertPptxToPng: (
    input: Uint8Array,
    options: { slides?: number[]; skipSystemFonts: true },
  ) => Promise<PngConversionReport>;
};

const defaultDependencies: PreviewPptxDependencies = {
  readFile,
  convertPptxToPng,
};

export async function previewPptx(
  input: PreviewPptxInput,
  dependencies: PreviewPptxDependencies = defaultDependencies,
): Promise<CallToolResult> {
  const slideInputError = validateSlideNumbers(input.slides);
  if (slideInputError !== undefined) {
    return errorResult(slideInputError);
  }

  let pptx: Uint8Array;
  try {
    pptx = await dependencies.readFile(input.filePath);
  } catch (error) {
    if (isFileNotFoundError(error)) {
      return errorResult(`PPTX file not found: ${input.filePath}`);
    }
    return errorResult(`Unable to read PPTX file '${input.filePath}': ${errorMessage(error)}`);
  }

  let report: PngConversionReport;
  try {
    report = await dependencies.convertPptxToPng(pptx, {
      ...(input.slides === undefined ? {} : { slides: input.slides }),
      skipSystemFonts: true,
    });
  } catch (error) {
    return errorResult(`Failed to convert PPTX file '${input.filePath}': ${errorMessage(error)}`);
  }

  const missingSlideNumbers = findMissingSlideNumbers(input.slides, report);
  if (missingSlideNumbers.length > 0) {
    return errorResult(
      `Slide number${missingSlideNumbers.length === 1 ? "" : "s"} out of range: ${missingSlideNumbers.join(", ")}`,
    );
  }

  const content = report.slides.map((slide) => ({
    type: "image" as const,
    data: Buffer.from(slide.png).toString("base64"),
    mimeType: "image/png",
  }));
  const structuredContent: PreviewPptxStructuredContent = {
    slides: report.slides.map((slide, contentIndex) => ({
      slideNumber: slide.slideNumber,
      contentIndex,
      width: slide.width,
      height: slide.height,
    })),
    diagnostics: summarizeDiagnostics(report),
    supportCoverage: report.supportCoverage,
  };

  return {
    content,
    structuredContent: { ...structuredContent },
  };
}

function validateSlideNumbers(slides: number[] | undefined): string | undefined {
  if (slides === undefined) return undefined;
  if (slides.length === 0) {
    return "slides must contain at least one 1-based slide number when provided";
  }

  const invalid = slides.filter((slideNumber) => !Number.isInteger(slideNumber) || slideNumber < 1);
  if (invalid.length > 0) {
    return `Invalid slide number${invalid.length === 1 ? "" : "s"}: ${invalid.join(", ")}. Slide numbers must be positive 1-based integers.`;
  }
  return undefined;
}

function findMissingSlideNumbers(
  requestedSlides: number[] | undefined,
  report: PngConversionReport,
): number[] {
  if (requestedSlides === undefined) return [];
  const renderedSlides = new Set(report.slides.map((slide) => slide.slideNumber));
  return [...new Set(requestedSlides)].filter((slideNumber) => !renderedSlides.has(slideNumber));
}

function summarizeDiagnostics(report: PngConversionReport) {
  const count = (severity: "info" | "warning" | "error") =>
    report.diagnostics.filter((diagnostic) => diagnostic.severity === severity).length;
  return {
    total: report.diagnostics.length,
    info: count("info"),
    warnings: count("warning"),
    errors: count("error"),
  };
}

function errorResult(message: string): CallToolResult {
  return {
    content: [{ type: "text", text: message }],
    isError: true,
  };
}

function isFileNotFoundError(error: unknown): boolean {
  return (
    error instanceof Error &&
    "code" in error &&
    typeof error.code === "string" &&
    error.code === "ENOENT"
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
