/**
 * PPTX から使用フォント名を収集する API。
 */
import {
  type CleanDocSource,
  type ComputedElement,
  type ComputedTextBody,
  createComputedView,
  readPptx,
  type SourceThemeFontScheme,
} from "../../../pptx-glimpse-document/src/experimental.js";

/** フォント収集結果 */
export interface UsedFonts {
  /** テーマで定義されたフォント */
  theme: {
    majorFont: string;
    minorFont: string;
    majorFontEa: string | null;
    minorFontEa: string | null;
    majorFontCs: string | null;
    minorFontCs: string | null;
  };
  /** テキストランやテーマで使用されているフォント名の一覧（重複なし、ソート済み） */
  fonts: string[];
}

const DEFAULT_THEME_FONTS: Required<Pick<SourceThemeFontScheme, "majorLatin" | "minorLatin">> &
  SourceThemeFontScheme = {
  majorLatin: "Calibri",
  minorLatin: "Calibri",
};

/**
 * PPTX をパースして使用されているフォント名を収集する。
 * レンダリングは行わないため軽量。
 */
export function collectUsedFonts(input: Buffer | Uint8Array): UsedFonts {
  const source = readPptx(input);
  const fontScheme = findThemeFontScheme(source);
  const computed = createComputedView(source);

  const fonts = new Set<string>();

  collectThemeFonts(fontScheme, fonts);

  const visitedTemplateParts = new Set<string>();
  for (const slide of computed.slides) {
    const templateElementsByPart = new Map<string, ComputedElement[]>();
    for (const element of slide.elements) {
      if (element.sourceLayer === "slide") {
        collectFontsFromElements([element], fonts, fontScheme);
        continue;
      }

      const key = `${element.sourceLayer}:${element.sourcePartPath}`;
      const elements = templateElementsByPart.get(key);
      if (elements) {
        elements.push(element);
      } else {
        templateElementsByPart.set(key, [element]);
      }
    }

    for (const [key, elements] of templateElementsByPart) {
      if (visitedTemplateParts.has(key)) continue;
      visitedTemplateParts.add(key);
      collectFontsFromElements(elements, fonts, fontScheme);
    }
  }

  return {
    theme: {
      majorFont: fontScheme.majorLatin,
      minorFont: fontScheme.minorLatin,
      majorFontEa: fontScheme.majorEastAsian ?? null,
      minorFontEa: fontScheme.minorEastAsian ?? null,
      majorFontCs: fontScheme.majorComplexScript ?? null,
      minorFontCs: fontScheme.minorComplexScript ?? null,
    },
    fonts: [...fonts].sort(),
  };
}

function findThemeFontScheme(
  source: CleanDocSource,
): Required<Pick<SourceThemeFontScheme, "majorLatin" | "minorLatin">> & SourceThemeFontScheme {
  const firstThemePartPath = source.slideMasters.find(
    (master) => master.themePartPath !== undefined,
  )?.themePartPath;
  const scheme =
    source.themes.find((theme) => theme.partPath === firstThemePartPath)?.fontScheme ??
    source.themes[0]?.fontScheme;
  return {
    ...DEFAULT_THEME_FONTS,
    ...scheme,
  };
}

function collectThemeFonts(fontScheme: SourceThemeFontScheme, fonts: Set<string>): void {
  addFont(fonts, fontScheme.majorLatin, fontScheme);
  addFont(fonts, fontScheme.minorLatin, fontScheme);
  addFont(fonts, fontScheme.majorEastAsian, fontScheme);
  addFont(fonts, fontScheme.minorEastAsian, fontScheme);
  addFont(fonts, fontScheme.majorComplexScript, fontScheme);
  addFont(fonts, fontScheme.minorComplexScript, fontScheme);
  addFont(fonts, fontScheme.majorJapanese, fontScheme);
  addFont(fonts, fontScheme.minorJapanese, fontScheme);
}

function collectFontsFromElements(
  elements: readonly ComputedElement[],
  fonts: Set<string>,
  fontScheme: SourceThemeFontScheme,
): void {
  for (const el of elements) {
    switch (el.kind) {
      case "shape":
        if (el.textBody) collectFontsFromTextBody(el.textBody, fonts, fontScheme);
        break;
      case "group":
        collectFontsFromElements(el.children, fonts, fontScheme);
        break;
      case "table":
        for (const row of el.table.rows) {
          for (const cell of row.cells) {
            if (cell.textBody) collectFontsFromTextBody(cell.textBody, fonts, fontScheme);
          }
        }
        break;
    }
  }
}

function collectFontsFromTextBody(
  textBody: ComputedTextBody,
  fonts: Set<string>,
  fontScheme: SourceThemeFontScheme,
): void {
  for (const para of textBody.paragraphs) {
    addFont(fonts, para.properties?.bulletFont, fontScheme);
    for (const run of para.runs) {
      addFont(fonts, run.properties?.typeface, fontScheme);
      addFont(fonts, run.properties?.typefaceEa, fontScheme);
      addFont(fonts, run.properties?.typefaceCs, fontScheme);
    }
  }
}

function addFont(
  fonts: Set<string>,
  font: string | undefined,
  fontScheme: SourceThemeFontScheme,
): void {
  const resolved = resolveThemeFontAlias(font, fontScheme);
  if (resolved) fonts.add(resolved);
}

function resolveThemeFontAlias(
  font: string | undefined,
  fontScheme: SourceThemeFontScheme,
): string | undefined {
  switch (font) {
    case "+mj-lt":
      return fontScheme.majorLatin;
    case "+mn-lt":
      return fontScheme.minorLatin;
    case "+mj-ea":
      return fontScheme.majorEastAsian ?? fontScheme.majorJapanese;
    case "+mn-ea":
      return fontScheme.minorEastAsian ?? fontScheme.minorJapanese;
    case "+mj-cs":
      return fontScheme.majorComplexScript;
    case "+mn-cs":
      return fontScheme.minorComplexScript;
    default:
      return font?.startsWith("+mj-") || font?.startsWith("+mn-") ? undefined : font;
  }
}
