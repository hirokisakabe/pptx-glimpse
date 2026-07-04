import type {
  ComputedElement,
  ComputedSlide,
  Diagnostic,
  PptxComputedView,
  SourceHandle,
} from "@pptx-glimpse/document";
import { createComputedView, readPptx } from "@pptx-glimpse/document";
import type {
  FontBuffer,
  FontMapping,
  LogLevel,
  OpentypeSetup,
  Slide,
  SlideElement,
  WarningEntry,
} from "@pptx-glimpse/renderer";
import {
  buildFontFaceStyle,
  createFontMapping,
  createOpentypeSetupFromBuffers,
  flushWarnings,
  FontUsageCollector,
  getWarningEntries,
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
   * Defaults to `"off"`. Diagnostics are always collected in the conversion
   * report; this option only controls console output. Use `"warn"` to print
   * summaries, or `"debug"` to print individual warning entries as they are
   * recorded.
   */
  logLevel?: LogLevel;
  /**
   * Additional directories that are scanned for font files.
   *
   * These directories are searched in addition to system font directories unless
   * `skipSystemFonts` is true. This option uses Node.js filesystem APIs; use
   * `fonts` when rendering in browsers, Edge Runtime, or other environments
   * where filesystem access is unavailable.
   */
  fontDirs?: string[];
  /**
   * Font file bytes supplied directly by the caller.
   *
   * Use this browser-safe option when fonts are fetched from a URL, bundled by
   * an application, or otherwise loaded without Node.js filesystem APIs. When
   * provided, these font buffers are used instead of system font scanning.
   */
  fonts?: FontBuffer[];
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

export interface ConversionDiagnostic {
  readonly source: "document" | "computed-view" | "renderer-adapter" | "renderer";
  readonly severity: "info" | "warning" | "error";
  readonly code: string;
  readonly message: string;
  readonly slideNumber?: number;
  readonly sourcePartPath?: string;
  readonly context?: string;
  readonly handle?: SourceHandle;
}

export interface SupportCoverageCounts {
  /**
   * Number of PPTX source/computed elements considered for rendering.
   */
  readonly inputElements: number;
  /**
   * Number of renderer-model elements produced for SVG / PNG output.
   */
  readonly outputElements: number;
  /**
   * Elements or raw nodes skipped because they are outside the supported render subset.
   */
  readonly skippedElements: number;
  /**
   * Elements skipped because a referenced PPTX part or relationship could not be resolved.
   */
  readonly unresolvedElements: number;
  /**
   * Elements or properties rendered with a fallback or ignored unsupported property.
   */
  readonly fallbackElements: number;
  /**
   * Diagnostics with warning severity that affect support/renderability confidence.
   */
  readonly warnings: number;
}

export interface SlideSupportCoverage extends SupportCoverageCounts {
  readonly slideNumber: number;
}

export interface SupportCoverage {
  /**
   * Support/renderability coverage summary. This is not a visual-match or pixel accuracy metric.
   */
  readonly overall: SupportCoverageCounts;
  readonly slides: readonly SlideSupportCoverage[];
}

export interface SvgConversionReport {
  readonly slides: readonly SlideSvg[];
  readonly diagnostics: readonly ConversionDiagnostic[];
  readonly supportCoverage: SupportCoverage;
}

export type SystemFontSetupLoader = (
  options: ConvertOptions | undefined,
) => Promise<OpentypeSetup | null>;

/**
 * Convert a PPTX file to SVG documents.
 *
 * @param input PPTX binary data.
 * @param options Conversion options. `slides` uses 1-based slide numbers; when
 * omitted, all slides are converted.
 * @returns A conversion report containing converted slides, diagnostics, and support coverage.
 *
 * Text is emitted as SVG paths by default for portable rendering. Set
 * `textOutput: "text"` to emit native `<text>` elements with embedded subset
 * fonts for inline browser SVG use. Font directories and font mapping options
 * control how PPTX font names are resolved for text measurement and text output.
 */
export async function convertPptxToSvg(
  input: Uint8Array,
  options?: ConvertOptions,
  loadSystemFontSetup?: SystemFontSetupLoader,
): Promise<SvgConversionReport> {
  const textOutput = options?.textOutput ?? "path";
  const logLevel = options?.logLevel ?? "off";
  const setup = await createOpentypeSetup(options, loadSystemFontSetup);
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
    initWarningLogger(logLevel === "off" ? "warn" : logLevel);

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

    const rendererWarningEntries = [...getWarningEntries()];
    if (logLevel === "off") {
      initWarningLogger("off");
    } else {
      flushWarnings();
    }

    const diagnostics: ConversionDiagnostic[] = [
      ...normalizeDocumentDiagnostics(source.diagnostics),
      ...collectSmartArtComputedViewDiagnostics(computed),
      ...normalizeRendererAdapterDiagnostics(adapted.diagnostics),
      ...normalizeRendererWarningDiagnostics(rendererWarningEntries),
    ];
    const supportCoverage = buildSupportCoverage(computed, adapted.slides, diagnostics);

    return { slides, diagnostics, supportCoverage };
  } finally {
    if (logLevel === "off") {
      initWarningLogger("off");
    }
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

function normalizeDocumentDiagnostics(diagnostics: readonly Diagnostic[]): ConversionDiagnostic[] {
  return diagnostics.map((diagnostic) => ({
    source: "document",
    severity: diagnostic.severity,
    code: diagnostic.code,
    message: diagnostic.message,
    ...(diagnostic.handle !== undefined ? { handle: diagnostic.handle } : {}),
    ...(diagnostic.handle?.partPath !== undefined
      ? { sourcePartPath: diagnostic.handle.partPath }
      : {}),
  }));
}

function collectSmartArtComputedViewDiagnostics(
  computed: PptxComputedView,
): ConversionDiagnostic[] {
  const diagnostics: ConversionDiagnostic[] = [];
  for (const slide of computed.slides) {
    for (const element of flattenComputedElements(slide.elements)) {
      if (element.kind !== "smartArt" || element.diagramDrawing === undefined) continue;
      for (const diagnostic of element.diagramDrawing.diagnostics) {
        diagnostics.push({
          source: "computed-view",
          severity: diagnostic.severity,
          code: diagnostic.code,
          message: diagnostic.message,
          slideNumber: slide.slideNumber,
          sourcePartPath: diagnostic.sourcePartPath,
        });
      }
    }
  }
  return diagnostics;
}

function normalizeRendererAdapterDiagnostics(
  diagnostics: readonly RendererAdapterDiagnostic[],
): ConversionDiagnostic[] {
  return diagnostics.map((diagnostic) => ({
    source: "renderer-adapter",
    severity: diagnostic.severity,
    code: diagnostic.code,
    message: diagnostic.message,
    ...(diagnostic.slideNumber !== undefined ? { slideNumber: diagnostic.slideNumber } : {}),
    ...(diagnostic.sourcePartPath !== undefined
      ? { sourcePartPath: diagnostic.sourcePartPath }
      : {}),
  }));
}

function normalizeRendererWarningDiagnostics(
  warnings: readonly WarningEntry[],
): ConversionDiagnostic[] {
  return warnings.map((warning) => ({
    source: "renderer",
    severity: "warning",
    code: `renderer.${warning.feature}`,
    message: warning.message,
    ...(warning.context !== undefined ? { context: warning.context } : {}),
    ...slideNumberFromWarningContext(warning.context),
  }));
}

function slideNumberFromWarningContext(context: string | undefined): { slideNumber?: number } {
  const match = /^Slide\s+(\d+)$/.exec(context ?? "");
  return match === null ? {} : { slideNumber: Number(match[1]) };
}

function buildSupportCoverage(
  computed: PptxComputedView,
  renderedSlides: readonly Slide[],
  diagnostics: readonly ConversionDiagnostic[],
): SupportCoverage {
  const renderedBySlide = new Map(renderedSlides.map((slide) => [slide.slideNumber, slide]));
  const slides = computed.slides.map((slide) => {
    const counts = buildSlideSupportCoverage(
      slide,
      renderedBySlide.get(slide.slideNumber),
      diagnostics,
    );
    return { slideNumber: slide.slideNumber, ...counts };
  });
  const slideTotals = slides.reduce<SupportCoverageCounts>(
    (total, slide) => ({
      inputElements: total.inputElements + slide.inputElements,
      outputElements: total.outputElements + slide.outputElements,
      skippedElements: total.skippedElements + slide.skippedElements,
      unresolvedElements: total.unresolvedElements + slide.unresolvedElements,
      fallbackElements: total.fallbackElements + slide.fallbackElements,
      warnings: total.warnings + slide.warnings,
    }),
    emptySupportCoverageCounts(),
  );

  return {
    overall: {
      ...slideTotals,
      warnings: diagnostics.filter((diagnostic) => diagnostic.severity === "warning").length,
    },
    slides,
  };
}

function buildSlideSupportCoverage(
  computedSlide: ComputedSlide,
  renderedSlide: Slide | undefined,
  diagnostics: readonly ConversionDiagnostic[],
): SupportCoverageCounts {
  const slideDiagnostics = diagnostics.filter(
    (diagnostic) => diagnostic.slideNumber === computedSlide.slideNumber,
  );
  return {
    inputElements: countComputedElements(computedSlide.elements),
    outputElements: renderedSlide !== undefined ? countRenderedElements(renderedSlide.elements) : 0,
    skippedElements: countDiagnosticsByCode(slideDiagnostics, isSkippedDiagnosticCode),
    unresolvedElements: countDiagnosticsByCode(slideDiagnostics, isUnresolvedDiagnosticCode),
    fallbackElements: countDiagnosticsByCode(slideDiagnostics, isFallbackDiagnosticCode),
    warnings: slideDiagnostics.filter((diagnostic) => diagnostic.severity === "warning").length,
  };
}

function emptySupportCoverageCounts(): SupportCoverageCounts {
  return {
    inputElements: 0,
    outputElements: 0,
    skippedElements: 0,
    unresolvedElements: 0,
    fallbackElements: 0,
    warnings: 0,
  };
}

function countDiagnosticsByCode(
  diagnostics: readonly ConversionDiagnostic[],
  predicate: (code: string) => boolean,
): number {
  return diagnostics.filter((diagnostic) => predicate(diagnostic.code)).length;
}

function isSkippedDiagnosticCode(code: string): boolean {
  return code === "pptx-computed-view-adapter.raw-element-skipped";
}

function isUnresolvedDiagnosticCode(code: string): boolean {
  return (
    code === "pptx-computed-view-adapter.unresolved-chart-skipped" ||
    code === "pptx-computed-view-adapter.unresolved-smartart-skipped" ||
    code === "pptx-computed-view-adapter.unresolved-image-skipped"
  );
}

function isFallbackDiagnosticCode(code: string): boolean {
  return (
    code === "pptx-computed-view-adapter.missing-transform" ||
    code === "pptx-computed-view-adapter.raw-background-ignored" ||
    code === "pptx-computed-view-adapter.raw-fill-ignored"
  );
}

function countComputedElements(elements: readonly ComputedElement[]): number {
  return flattenComputedElements(elements).length;
}

function flattenComputedElements(elements: readonly ComputedElement[]): ComputedElement[] {
  const flattened: ComputedElement[] = [];
  for (const element of elements) {
    flattened.push(element);
    if (element.kind === "group") {
      flattened.push(...flattenComputedElements(element.children));
    }
    if (element.kind === "smartArt" && element.diagramDrawing !== undefined) {
      flattened.push(...flattenComputedElements(element.diagramDrawing.children));
    }
  }
  return flattened;
}

function countRenderedElements(elements: readonly SlideElement[]): number {
  let count = 0;
  for (const element of elements) {
    count++;
    if (element.type === "group") {
      count += countRenderedElements(element.children);
    }
  }
  return count;
}

function injectIntoSvgDefs(svg: string, content: string): string {
  const openTagEnd = svg.indexOf(">");
  if (openTagEnd === -1) return svg;
  return `${svg.slice(0, openTagEnd + 1)}<defs>${content}</defs>${svg.slice(openTagEnd + 1)}`;
}

async function createOpentypeSetup(
  options: ConvertOptions | undefined,
  loadSystemFontSetup: SystemFontSetupLoader | undefined,
) {
  if (options?.fonts !== undefined) {
    return createOpentypeSetupFromBuffers(options.fonts, options.fontMapping);
  }
  if (!shouldLoadSystemFonts(options) || loadSystemFontSetup === undefined) {
    return null;
  }

  return loadSystemFontSetup(options);
}

function shouldLoadSystemFonts(options: ConvertOptions | undefined): boolean {
  return options?.skipSystemFonts !== true || (options?.fontDirs?.length ?? 0) > 0;
}
