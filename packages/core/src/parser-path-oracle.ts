import { readFileSync, statSync } from "node:fs";

import type { ShapeElement, SlideElement } from "@pptx-glimpse/renderer";
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
import { clearXmlCache, enableXmlCache } from "./parser/xml-parser.js";
import type { ParsedSlide } from "./pptx-data-parser.js";
import { parsePptxData, parseSlideWithLayout } from "./pptx-data-parser.js";

/**
 * Explicit old-parser render oracle for document-path parity checks.
 *
 * Public conversion defaults to the PptxSourceModel document path. Keep this module
 * internal so old parser semantics do not leak back into core orchestration.
 */
async function convertPptxToSvgViaParserPath(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideSvg[]> {
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
  enableXmlCache();
  try {
    initWarningLogger(options?.logLevel ?? "off");

    const data = parsePptxData(input);
    setScriptFonts(data.theme.fontScheme.majorFontJpan, data.theme.fontScheme.minorFontJpan);

    const targetSlides = options?.slides
      ? data.slidePaths.filter((s) => options.slides!.includes(s.slideNumber))
      : data.slidePaths;

    if (data.slidePaths.length === 0) {
      warn("presentation.noSlides", "No slides found in the PPTX file");
    }

    const results: SlideSvg[] = [];
    for (const { slideNumber, path } of targetSlides) {
      const parsed = parseSlideWithLayout(slideNumber, path, data);
      if (!parsed) continue;

      const { slide } = parsed;
      slide.elements = buildEffectiveSlideElements(parsed);

      fontUsageCollector?.reset();
      let svg = renderSlideToSvg(slide, data.presInfo.slideSize);
      if (fontUsageCollector && setup) {
        const style = await buildFontFaceStyle(fontUsageCollector.getUsages(), setup.fontResolver);
        if (style) {
          svg = injectIntoSvgDefs(svg, style);
        }
      }
      results.push({ slideNumber, svg });
    }

    flushWarnings();
    return results;
  } finally {
    clearXmlCache();
    resetTextMeasurer();
    resetTextPathFontResolver();
    resetFontUsageCollector();
    resetFontMapping();
    resetScriptFonts();
  }
}

export async function convertPptxToPngViaParserPath(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  return runParserPathExclusive(() => convertPptxToPngViaParserPathUnsafe(input, options));
}

async function convertPptxToPngViaParserPathUnsafe(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  const svgResults = await convertPptxToSvgViaParserPath(input, { ...options, textOutput: "path" });
  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;
  const fontBuffers = loadFontBuffers(options?.fontDirs, options?.skipSystemFonts);

  const results: SlideImage[] = [];
  for (const { slideNumber, svg } of svgResults) {
    const pngResult = await svgToPng(svg, { width, height, fontBuffers });
    results.push({
      slideNumber,
      png: pngResult.png,
      width: pngResult.width,
      height: pngResult.height,
    });
  }

  return results;
}

let parserPathQueue: Promise<void> = Promise.resolve();

async function runParserPathExclusive<T>(fn: () => Promise<T>): Promise<T> {
  const previous = parserPathQueue;
  let release!: () => void;
  parserPathQueue = new Promise((resolve) => {
    release = resolve;
  });

  await previous;
  try {
    return await fn();
  } finally {
    release();
  }
}

export function buildEffectiveSlideElements(parsed: ParsedSlide): SlideElement[] {
  const effectiveMasterElements =
    parsed.slide.showMasterSp && parsed.layoutShowMasterSp ? parsed.masterElements : [];
  return mergeElements(effectiveMasterElements, parsed.layoutElements, parsed.slide.elements);
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
  const key = `${[...(fontDirs ?? [])].sort().join("\0")}\n${skipSystemFonts ?? false}`;
  if (cachedFontBuffers !== null && cachedFontBuffersKey === key) {
    return cachedFontBuffers;
  }

  const allPaths = collectFontFilePaths(fontDirs, skipSystemFonts);
  const ttfOtfPaths = allPaths.filter((p) => {
    const lower = p.toLowerCase();
    return lower.endsWith(".ttf") || lower.endsWith(".otf");
  });

  const pathsWithSize: { path: string; size: number }[] = [];
  for (const p of ttfOtfPaths) {
    try {
      pathsWithSize.push({ path: p, size: statSync(p).size });
    } catch {
      // Ignore unreadable font files.
    }
  }
  pathsWithSize.sort((a, b) => a.size - b.size);

  const buffers: Uint8Array[] = [];
  let totalSize = 0;
  for (const { path, size } of pathsWithSize) {
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

function mergeElements(
  masterElements: SlideElement[],
  layoutElements: SlideElement[],
  slideElements: SlideElement[],
): SlideElement[] {
  const filterTemplatePlaceholders = (elements: SlideElement[]) =>
    elements.filter((el) => {
      if (el.type !== "shape") return true;
      return !el.placeholderType;
    });

  const filterEmptySlidePlaceholders = (elements: SlideElement[]) =>
    elements.filter((el) => !(el.type === "shape" && isEmptyPlaceholder(el)));

  const filteredMaster = filterTemplatePlaceholders(masterElements);
  const filteredLayout = filterTemplatePlaceholders(layoutElements);
  const filteredSlide = filterEmptySlidePlaceholders(slideElements);

  return [...filteredMaster, ...filteredLayout, ...filteredSlide];
}

function isEmptyPlaceholder(shape: ShapeElement): boolean {
  if (!shape.placeholderType) return false;
  const paragraphs = shape.textBody?.paragraphs;
  if (!paragraphs || paragraphs.length === 0) return true;
  return !paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
}
