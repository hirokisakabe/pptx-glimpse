/**
 * PPTX から使用フォント名を収集する API。
 */
import type { SlideElement } from "./model/shape.js";
import type { TextBody } from "./model/text.js";
import type { FontScheme } from "./model/theme.js";
import { parsePptxData, parseSlideWithLayout } from "./pptx-data-parser.js";

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

/**
 * PPTX をパースして使用されているフォント名を収集する。
 * レンダリングは行わないため軽量。
 */
export async function collectUsedFonts(input: Buffer | Uint8Array): Promise<UsedFonts> {
  const data = await parsePptxData(input);
  const fontScheme = data.theme.fontScheme;

  const fonts = new Set<string>();

  // テーマフォントを収集
  collectThemeFonts(fontScheme, fonts);

  // 各スライドから収集
  for (const { slideNumber, path } of data.slidePaths) {
    const parsed = parseSlideWithLayout(slideNumber, path, data);
    if (!parsed) continue;
    collectFontsFromElements(parsed.slide.elements, fonts);
  }

  // マスター要素からも収集
  collectFontsFromElements(data.masterElements, fonts);

  return {
    theme: {
      majorFont: fontScheme.majorFont,
      minorFont: fontScheme.minorFont,
      majorFontEa: fontScheme.majorFontEa,
      minorFontEa: fontScheme.minorFontEa,
      majorFontCs: fontScheme.majorFontCs,
      minorFontCs: fontScheme.minorFontCs,
    },
    fonts: [...fonts].sort(),
  };
}

function collectThemeFonts(fontScheme: FontScheme, fonts: Set<string>): void {
  fonts.add(fontScheme.majorFont);
  fonts.add(fontScheme.minorFont);
  if (fontScheme.majorFontEa) fonts.add(fontScheme.majorFontEa);
  if (fontScheme.minorFontEa) fonts.add(fontScheme.minorFontEa);
  if (fontScheme.majorFontCs) fonts.add(fontScheme.majorFontCs);
  if (fontScheme.minorFontCs) fonts.add(fontScheme.minorFontCs);
}

function collectFontsFromElements(elements: SlideElement[], fonts: Set<string>): void {
  for (const el of elements) {
    switch (el.type) {
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

function collectFontsFromTextBody(textBody: TextBody, fonts: Set<string>): void {
  for (const para of textBody.paragraphs) {
    if (para.properties.bulletFont) fonts.add(para.properties.bulletFont);
    for (const run of para.runs) {
      if (run.properties.fontFamily) fonts.add(run.properties.fontFamily);
      if (run.properties.fontFamilyEa) fonts.add(run.properties.fontFamilyEa);
      if (run.properties.fontFamilyCs) fonts.add(run.properties.fontFamilyCs);
    }
  }
}
