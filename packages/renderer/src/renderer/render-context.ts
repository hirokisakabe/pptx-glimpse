import { DEFAULT_FONT_MAPPING, type FontMapping, getMappedFont } from "../font/font-mapping.js";
import { getFontMapping } from "../font/font-mapping-context.js";
import type { FontUsageCollector } from "../font/font-usage-collector.js";
import { getFontUsageCollector } from "../font/font-usage-collector.js";
import { getScriptFonts } from "../font/script-font-context.js";
import { DefaultTextMeasurer, getTextMeasurer, type TextMeasurer } from "../font/text-measurer.js";
import type { TextPathFontResolver } from "../font/text-path-context.js";
import { getTextPathFontResolver } from "../font/text-path-context.js";
import {
  createWarningLogger,
  getActiveWarningLogger,
  type WarningLogger,
} from "../warning-logger.js";

export interface RendererScriptFonts {
  readonly majorJpan: string | null;
  readonly minorJpan: string | null;
}

export interface RendererContext {
  readonly textMeasurer: TextMeasurer;
  readonly textPathFontResolver: TextPathFontResolver | null;
  readonly fontMapping: FontMapping;
  readonly fontUsageCollector: FontUsageCollector | null;
  readonly scriptFonts: RendererScriptFonts;
  readonly warningLogger: WarningLogger;
}

export function createRendererContext(overrides: Partial<RendererContext> = {}): RendererContext {
  return {
    textMeasurer: overrides.textMeasurer ?? new DefaultTextMeasurer(),
    textPathFontResolver: overrides.textPathFontResolver ?? null,
    fontMapping: overrides.fontMapping ?? { ...DEFAULT_FONT_MAPPING },
    fontUsageCollector: overrides.fontUsageCollector ?? null,
    scriptFonts: overrides.scriptFonts ?? { majorJpan: null, minorJpan: null },
    warningLogger: overrides.warningLogger ?? createWarningLogger("off"),
  };
}

export function createLegacyRendererContext(): RendererContext {
  return {
    textMeasurer: getTextMeasurer(),
    textPathFontResolver: getTextPathFontResolver(),
    fontMapping: getFontMapping(),
    fontUsageCollector: getFontUsageCollector(),
    scriptFonts: getScriptFonts(),
    warningLogger: getActiveWarningLogger(),
  };
}

export function getJpanFallbackFontFromContext(context: RendererContext): string | null {
  return context.scriptFonts.majorJpan ?? context.scriptFonts.minorJpan;
}

export function getMappedFontFromContext(
  fontFamily: string | null | undefined,
  context: RendererContext,
): string | null {
  return getMappedFont(fontFamily, context.fontMapping);
}
