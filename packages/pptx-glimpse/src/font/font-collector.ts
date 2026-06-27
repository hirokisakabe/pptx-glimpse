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

  for (const slide of computed.slides) {
    collectFontsFromElements(slide.elements, fonts);
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
  addFont(fonts, fontScheme.majorLatin);
  addFont(fonts, fontScheme.minorLatin);
  addFont(fonts, fontScheme.majorEastAsian);
  addFont(fonts, fontScheme.minorEastAsian);
  addFont(fonts, fontScheme.majorComplexScript);
  addFont(fonts, fontScheme.minorComplexScript);
}

function collectFontsFromElements(elements: readonly ComputedElement[], fonts: Set<string>): void {
  for (const el of elements) {
    switch (el.kind) {
      case "shape":
        if (el.textBody) collectFontsFromTextBody(el.textBody, fonts);
        break;
      case "group":
        collectFontsFromElements(el.children, fonts);
        break;
      case "table":
        for (const row of el.table.rows) {
          for (const cell of row.cells) {
            if (cell.textBody) collectFontsFromTextBody(cell.textBody, fonts);
          }
        }
        break;
    }
  }
}

function collectFontsFromTextBody(textBody: ComputedTextBody, fonts: Set<string>): void {
  for (const para of textBody.paragraphs) {
    addFont(fonts, para.properties?.bulletFont);
    for (const run of para.runs) {
      addFont(fonts, run.properties?.typeface);
      addFont(fonts, run.properties?.typefaceEa);
      addFont(fonts, run.properties?.typefaceCs);
    }
  }
}

function addFont(fonts: Set<string>, font: string | undefined): void {
  if (font) fonts.add(font);
}
