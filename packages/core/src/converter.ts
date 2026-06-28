import { readFileSync, statSync } from "node:fs";

import { createComputedView, readPptx } from "@pptx-glimpse/document";
import type { FontMapping } from "@pptx-glimpse/renderer";
import type { LogLevel } from "@pptx-glimpse/renderer";
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

import {
  adaptComputedViewToRendererModel,
  type RendererAdapterDiagnostic,
} from "./pptx-computed-view-renderer-adapter.js";

/**
 * Options shared by PPTX-to-SVG and PPTX-to-PNG conversion.
 */
export interface ConvertOptions {
  /**
   * Target slide numbers to convert, using PowerPoint-style 1-based numbering.
   *
   * When omitted, every slide in the presentation is converted. Missing or
   * out-of-range slide numbers produce no output for those entries.
   */
  slides?: number[];
  /**
   * Output width in pixels.
   *
   * PNG output rasterizes to this width while preserving the slide aspect
   * ratio. Defaults to 960. SVG output keeps the slide's native pixel size from
   * the PPTX slide dimensions and does not use this option.
   */
  width?: number;
  /**
   * Output height in pixels.
   *
   * PNG rasterization currently always uses `width`, either the provided value
   * or the default width, so this option is ignored by the public conversion
   * APIs. SVG output keeps the slide's native pixel size and does not use this
   * option.
   */
  height?: number;
  /**
   * Warning log level for unsupported or approximated PPTX features.
   *
   * Defaults to `"off"`. Use `"warn"` to collect and print summaries, or
   * `"debug"` to print individual warning entries as they are recorded.
   */
  logLevel?: LogLevel;
  /**
   * Additional directories that are scanned for font files.
   *
   * These directories are searched in addition to system font directories unless
   * `skipSystemFonts` is true.
   */
  fontDirs?: string[];
  /**
   * Custom PPTX font name to replacement font name mapping.
   *
   * Entries are merged with `DEFAULT_FONT_MAPPING`; user-provided entries take
   * precedence. This is useful when a PPTX references proprietary or corporate
   * fonts that should render with installed alternatives.
   */
  fontMapping?: FontMapping;
  /**
   * Skip OS system font directories and only scan `fontDirs`.
   *
   * This is useful in containers or serverless environments where bundled fonts
   * should be the only fonts used.
   */
  skipSystemFonts?: boolean;
  /**
   * Text output mode for SVG conversion.
   *
   * Defaults to `"path"`, which emits glyph outlines as `<path>` elements and
   * does not require fonts in the SVG viewing environment. `"text"` emits native
   * `<text>` elements with subset-font `@font-face` data URIs, enabling browser
   * text selection and native text rendering for inline SVG.
   *
   * Embedded fonts and native text may not render as expected when the SVG is
   * loaded through `<img src="...svg">` or sanitized. `convertPptxToPng` ignores
   * this option and always renders with `"path"` output because resvg does not
   * interpret the embedded `@font-face` rules used by SVG text output.
   */
  textOutput?: "path" | "text";
}

/**
 * SVG conversion result for one slide.
 */
export interface SlideSvg {
  /**
   * Original slide number in the PPTX file, using 1-based numbering.
   */
  slideNumber: number;
  /**
   * Complete SVG document string for the slide.
   */
  svg: string;
}

/**
 * PNG conversion result for one slide.
 */
export interface SlideImage {
  /**
   * Original slide number in the PPTX file, using 1-based numbering.
   */
  slideNumber: number;
  /**
   * PNG image bytes for the rendered slide.
   */
  png: Buffer;
  /**
   * Actual output image width in pixels after rasterization.
   */
  width: number;
  /**
   * Actual output image height in pixels after rasterization.
   */
  height: number;
}

type ConversionDiagnostic = RendererAdapterDiagnostic;

interface SvgConversionResult {
  readonly slides: readonly SlideSvg[];
  readonly diagnostics: readonly ConversionDiagnostic[];
}

interface PngConversionResult {
  readonly slides: readonly SlideImage[];
  readonly diagnostics: readonly ConversionDiagnostic[];
}

/**
 * Convert a PPTX file to SVG documents.
 *
 * @param input PPTX binary data as a Node.js `Buffer` or `Uint8Array`.
 * @param options Conversion options. `slides` uses 1-based slide numbers; when
 * omitted, all slides are converted.
 * @returns One SVG result per converted slide, preserving original slide numbers.
 *
 * Text is emitted as SVG paths by default for portable rendering. Set
 * `textOutput: "text"` to emit native `<text>` elements with embedded subset
 * fonts for inline browser SVG use. Font directories and font mapping options
 * control how PPTX font names are resolved for text measurement and text output.
 */
export async function convertPptxToSvg(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideSvg[]> {
  const result = await convertPptxToSvgInternal(input, options);
  return [...result.slides];
}

/**
 * Convert a PPTX file to PNG images.
 *
 * @param input PPTX binary data as a Node.js `Buffer` or `Uint8Array`.
 * @param options Conversion options. `slides` uses 1-based slide numbers; when
 * omitted, all slides are converted.
 * @returns One PNG result per converted slide, preserving original slide numbers
 * and reporting the actual rasterized image size.
 *
 * PNG conversion first renders each slide to SVG and then rasterizes it with
 * resvg. The `textOutput` option is intentionally ignored: PNG rendering always
 * uses path-based text output because resvg does not interpret the embedded
 * `@font-face` rules used by SVG text mode. Font directories and font mapping
 * options are still used to resolve glyph outlines and text metrics.
 */
export async function convertPptxToPng(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  const result = await convertPptxToPngInternal(input, options);
  return [...result.slides];
}

async function convertPptxToSvgInternal(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SvgConversionResult> {
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
      throw new Error("Converter requires a computed slide size");
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

async function convertPptxToPngInternal(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<PngConversionResult> {
  const svgResult = await convertPptxToSvgInternal(input, {
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
