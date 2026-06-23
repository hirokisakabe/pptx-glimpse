import { readFileSync, statSync } from "node:fs";

import {
  buildFontFaceStyle,
  collectFontFilePaths,
  createFontMapping,
  createOpentypeSetupFromSystem,
  DEFAULT_OUTPUT_WIDTH,
  flushWarnings,
  FontUsageCollector,
  initWarningLogger,
  renderSlideToSvg,
  resetFontMapping,
  resetFontUsageCollector,
  resetScriptFonts,
  resetTextMeasurer,
  resetTextPathFontResolver,
  setFontMapping,
  setFontUsageCollector,
  setTextMeasurer,
  setTextPathFontResolver,
  svgToPng,
  warn,
} from "pptx-glimpse-renderer";

import {
  type CleanDocComputedView,
  createComputedView,
  readPptx,
} from "../../pptx-glimpse-document/src/experimental.js";
import {
  adaptComputedViewToRendererModel,
  type RendererAdapterDiagnostic,
} from "./cleandoc-renderer-adapter.js";
import type { ConvertOptions, SlideImage, SlideSvg } from "./converter.js";

type DocumentRenderDiagnostic = RendererAdapterDiagnostic | DocumentRenderPathDiagnostic;

interface DocumentRenderPathDiagnostic {
  readonly severity: "warning";
  readonly code: "document-render.cjk-font-context-unsupported";
  readonly message: string;
}

interface DocumentPathSvgResult {
  readonly slides: readonly SlideSvg[];
  readonly diagnostics: readonly DocumentRenderDiagnostic[];
}

interface DocumentPathPngResult {
  readonly slides: readonly SlideImage[];
  readonly diagnostics: readonly DocumentRenderDiagnostic[];
}

/**
 * Render orchestration for dogfooding the CleanDoc document path.
 */
export async function convertPptxToSvgViaDocumentPath(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<DocumentPathSvgResult> {
  const textOutput = options?.textOutput ?? "path";
  const setup = await createOpentypeSetupFromSystem(
    options?.fontDirs,
    options?.fontMapping,
    options?.skipSystemFonts,
  );
  if (setup) {
    setTextMeasurer(setup.measurer);
    if (textOutput !== "text") {
      setTextPathFontResolver(setup.fontResolver);
    }
  }

  const fontUsageCollector = textOutput === "text" ? new FontUsageCollector() : null;
  if (fontUsageCollector) {
    setFontUsageCollector(fontUsageCollector);
  }
  setFontMapping(createFontMapping(options?.fontMapping));

  try {
    initWarningLogger(options?.logLevel ?? "off");

    const source = readPptx(input);
    if (source.presentation.slidePartPaths.length === 0) {
      warn("presentation.noSlides", "No slides found in the PPTX file");
    }

    const computed = createComputedView(source, { slides: options?.slides });
    const adapted = adaptComputedViewToRendererModel(computed);
    const diagnostics = [...collectDocumentRenderDiagnostics(computed), ...adapted.diagnostics];
    emitDocumentRenderDiagnostics(diagnostics);
    const slideSize = adapted.slideSize;
    if (slideSize === undefined && adapted.slides.length > 0) {
      throw new Error("Document render path requires a computed slide size");
    }

    const slides: SlideSvg[] = [];
    for (const slide of adapted.slides) {
      if (slideSize === undefined) continue;
      fontUsageCollector?.reset();
      let svg = renderSlideToSvg(slide, slideSize);
      if (fontUsageCollector && setup) {
        const style = await buildFontFaceStyle(fontUsageCollector.getUsages(), setup.fontResolver);
        if (style) {
          svg = injectIntoSvgDefs(svg, style);
        }
      }
      slides.push({ slideNumber: slide.slideNumber, svg });
    }

    flushWarnings();
    return { slides, diagnostics };
  } finally {
    resetTextMeasurer();
    resetTextPathFontResolver();
    resetFontUsageCollector();
    resetFontMapping();
    resetScriptFonts();
  }
}

/**
 * Internal / experimental PNG path layered on top of the CleanDoc SVG path and
 * the existing renderer PNG conversion.
 */
export async function convertPptxToPngViaDocumentPath(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<DocumentPathPngResult> {
  const svgResult = await convertPptxToSvgViaDocumentPath(input, {
    ...options,
    textOutput: "path",
  });
  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;
  const fontBuffers = loadFontBuffers(options?.fontDirs, options?.skipSystemFonts);

  const slides: SlideImage[] = [];
  for (const { slideNumber, svg } of svgResult.slides) {
    const pngResult = await svgToPng(svg, { width, height, fontBuffers });
    slides.push({
      slideNumber,
      png: pngResult.png,
      width: pngResult.width,
      height: pngResult.height,
    });
  }

  return { slides, diagnostics: svgResult.diagnostics };
}

function collectDocumentRenderDiagnostics(
  computed: CleanDocComputedView,
): DocumentRenderPathDiagnostic[] {
  if (!containsCjkText(computed)) return [];
  return [
    {
      severity: "warning",
      code: "document-render.cjk-font-context-unsupported",
      message:
        "CleanDoc document render path does not yet expose East Asian or complex-script theme fonts; CJK text may not match the public converter output.",
    },
  ];
}

function emitDocumentRenderDiagnostics(diagnostics: readonly DocumentRenderDiagnostic[]): void {
  for (const diagnostic of diagnostics) {
    warn(diagnostic.code, diagnostic.message, getDiagnosticContext(diagnostic));
  }
}

function getDiagnosticContext(diagnostic: DocumentRenderDiagnostic): string | undefined {
  if (!("slideNumber" in diagnostic) || diagnostic.slideNumber === undefined) return undefined;
  if (diagnostic.sourcePartPath === undefined) return `Slide ${diagnostic.slideNumber}`;
  return `Slide ${diagnostic.slideNumber}, ${diagnostic.sourcePartPath}`;
}

function containsCjkText(computed: CleanDocComputedView): boolean {
  return computed.slides.some((slide) =>
    slide.elements.some(
      (element) =>
        element.kind === "shape" &&
        element.textBody?.paragraphs.some((paragraph) =>
          paragraph.runs.some((run) => cjkTextPattern.test(run.text)),
        ) === true,
    ),
  );
}

const cjkTextPattern = /[\u3040-\u30ff\u3400-\u9fff\uff00-\uffef]/;

function injectIntoSvgDefs(svg: string, content: string): string {
  const openTagEnd = svg.indexOf(">");
  if (openTagEnd === -1) return svg;
  return `${svg.slice(0, openTagEnd + 1)}<defs>${content}</defs>${svg.slice(openTagEnd + 1)}`;
}

let cachedFontBuffers: Uint8Array[] | null = null;
let cachedFontBuffersKey: string | null = null;

const MAX_TOTAL_FONT_BUFFER_BYTES = 100 * 1024 * 1024;

function loadFontBuffers(fontDirs?: string[], skipSystemFonts?: boolean): Uint8Array[] {
  const key = `${(fontDirs ?? []).join("\0")}\n${skipSystemFonts ?? false}`;
  if (cachedFontBuffers !== null && cachedFontBuffersKey === key) {
    return cachedFontBuffers;
  }

  const fontPaths = collectFontFilePaths(fontDirs, skipSystemFonts).filter((path) => {
    const lower = path.toLowerCase();
    return lower.endsWith(".ttf") || lower.endsWith(".otf");
  });
  const readableFontPaths: { path: string; size: number }[] = [];
  for (const path of fontPaths) {
    try {
      readableFontPaths.push({ path, size: statSync(path).size });
    } catch {
      // Ignore unreadable font files.
    }
  }
  readableFontPaths.sort((a, b) => a.size - b.size);

  const buffers: Uint8Array[] = [];
  let totalSize = 0;
  for (const { path, size } of readableFontPaths) {
    if (totalSize + size > MAX_TOTAL_FONT_BUFFER_BYTES) break;
    try {
      buffers.push(new Uint8Array(readFileSync(path)));
      totalSize += size;
    } catch {
      // Ignore fonts that disappear between stat and read.
    }
  }

  cachedFontBuffers = buffers;
  cachedFontBuffersKey = key;
  return buffers;
}
