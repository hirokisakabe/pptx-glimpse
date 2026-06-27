import { readFileSync, statSync } from "node:fs";

import { createComputedView, readPptx } from "@pptx-glimpse/document/experimental";
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
  setScriptFonts,
  setTextMeasurer,
  setTextPathFontResolver,
  svgToPng,
  warn,
} from "@pptx-glimpse/renderer";

import type { ConvertOptions, SlideImage, SlideSvg } from "./converter.js";
import {
  adaptComputedViewToRendererModel,
  type RendererAdapterDiagnostic,
} from "./pptx-computed-view-renderer-adapter.js";

type DocumentRenderDiagnostic = RendererAdapterDiagnostic;

interface DocumentPathSvgResult {
  readonly slides: readonly SlideSvg[];
  readonly diagnostics: readonly DocumentRenderDiagnostic[];
}

interface DocumentPathPngResult {
  readonly slides: readonly SlideImage[];
  readonly diagnostics: readonly DocumentRenderDiagnostic[];
}

/**
 * PptxSourceModel document-path render orchestration used by the public converter
 * default. Parser-path helpers remain available for parity checks.
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
    const scriptFontScheme = findScriptFontScheme(source);
    setScriptFonts(
      scriptFontScheme?.majorJapanese ?? null,
      scriptFontScheme?.minorJapanese ?? null,
    );
    if (source.presentation.slidePartPaths.length === 0) {
      warn("presentation.noSlides", "No slides found in the PPTX file");
    }

    const computed = createComputedView(source, { slides: options?.slides });
    const adapted = adaptComputedViewToRendererModel(computed);
    const diagnostics = adapted.diagnostics;
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

function findScriptFontScheme(source: ReturnType<typeof readPptx>) {
  const firstThemePartPath = source.slideMasters.find(
    (master) => master.themePartPath !== undefined,
  )?.themePartPath;
  return (
    source.themes.find((theme) => theme.partPath === firstThemePartPath)?.fontScheme ??
    source.themes[0]?.fontScheme
  );
}

/**
 * Internal / experimental PNG path layered on top of the PptxSourceModel SVG path and
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
  readableFontPaths.sort((a, b) => a.size - b.size || a.path.localeCompare(b.path));

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
