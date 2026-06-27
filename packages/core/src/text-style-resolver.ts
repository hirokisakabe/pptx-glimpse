import type { ShapeElement, SlideElement } from "@pptx-glimpse/renderer";
import type {
  DefaultRunProperties,
  DefaultTextStyle,
  PlaceholderStyleInfo,
  TxStyles,
} from "@pptx-glimpse/renderer";
import type { FontScheme } from "@pptx-glimpse/renderer";

import { resolveThemeFont } from "./parser/text-style-parser.js";

export interface TextStyleContext {
  layoutPlaceholderStyles: PlaceholderStyleInfo[];
  masterPlaceholderStyles: PlaceholderStyleInfo[];
  txStyles?: TxStyles;
  defaultTextStyle?: DefaultTextStyle;
  fontScheme?: FontScheme | null;
}

/**
 * スライド要素に対してテキストスタイル継承チェーンを適用する。
 * 各 RunProperties の null 値（fontSize, fontFamily, fontFamilyEa）を
 * 継承チェーンから解決して埋める。
 */
export function applyTextStyleInheritance(
  elements: SlideElement[],
  context: TextStyleContext,
): void {
  for (const element of elements) {
    if (element.type === "shape") {
      resolveShapeTextInheritance(element, context);
    } else if (element.type === "group") {
      applyTextStyleInheritance(element.children, context);
    }
  }
}

function resolveShapeTextInheritance(shape: ShapeElement, context: TextStyleContext): void {
  if (!shape.textBody) return;

  const layoutLstStyle = findMatchingPlaceholderStyle(
    shape.placeholderType,
    shape.placeholderIdx,
    context.layoutPlaceholderStyles,
  );
  const masterLstStyle = findMatchingPlaceholderStyle(
    shape.placeholderType,
    shape.placeholderIdx,
    context.masterPlaceholderStyles,
  );
  const txStyle = getTxStyleForPlaceholder(shape.placeholderType, context.txStyles);

  const chainSources = [layoutLstStyle, masterLstStyle, txStyle, context.defaultTextStyle];

  for (const paragraph of shape.textBody.paragraphs) {
    const level = paragraph.properties.level;

    // 段落 alignment の継承解決
    if (paragraph.properties.alignment === null) {
      for (const source of chainSources) {
        if (!source) continue;
        const alignment = source.levels[level]?.alignment ?? source.defaultParagraph?.alignment;
        if (alignment) {
          paragraph.properties.alignment = alignment;
          break;
        }
      }
      // 継承チェーンでも見つからなければ左揃えにフォールバック
      if (paragraph.properties.alignment === null) {
        paragraph.properties.alignment = "l";
      }
    }

    // 段落 bullet の継承解決
    if (paragraph.properties.bullet === null) {
      for (const source of chainSources) {
        if (!source) continue;
        const levelProps = source.levels[level] ?? source.defaultParagraph;
        if (!levelProps?.bullet) continue;
        paragraph.properties.bullet = levelProps.bullet;
        break;
      }
    }

    // bullet 補助属性の継承解決（bullet 本体とは独立して解決）
    const hasBullet =
      paragraph.properties.bullet !== null && paragraph.properties.bullet.type !== "none";
    if (
      hasBullet &&
      (paragraph.properties.bulletFont === null ||
        paragraph.properties.bulletColor === null ||
        paragraph.properties.bulletSizePct === null)
    ) {
      for (const source of chainSources) {
        if (!source) continue;
        const levelProps = source.levels[level] ?? source.defaultParagraph;
        if (!levelProps) continue;
        if (paragraph.properties.bulletFont === null && levelProps.bulletFont) {
          paragraph.properties.bulletFont = levelProps.bulletFont;
        }
        if (paragraph.properties.bulletColor === null && levelProps.bulletColor) {
          paragraph.properties.bulletColor = levelProps.bulletColor;
        }
        if (paragraph.properties.bulletSizePct === null && levelProps.bulletSizePct) {
          paragraph.properties.bulletSizePct = levelProps.bulletSizePct;
        }
        if (
          paragraph.properties.bulletFont !== null &&
          paragraph.properties.bulletColor !== null &&
          paragraph.properties.bulletSizePct !== null
        ) {
          break;
        }
      }
    }

    // 段落 marginLeft / indent の継承解決
    if (paragraph.properties.marginLeft === null || paragraph.properties.indent === null) {
      for (const source of chainSources) {
        if (!source) continue;
        const levelProps = source.levels[level] ?? source.defaultParagraph;
        if (!levelProps) continue;
        if (paragraph.properties.marginLeft === null && levelProps.marginLeft !== undefined) {
          paragraph.properties.marginLeft = levelProps.marginLeft;
        }
        if (paragraph.properties.indent === null && levelProps.indent !== undefined) {
          paragraph.properties.indent = levelProps.indent;
        }
        if (paragraph.properties.marginLeft !== null && paragraph.properties.indent !== null) {
          break;
        }
      }
    }

    for (const run of paragraph.runs) {
      const props = run.properties;

      for (const source of chainSources) {
        if (
          props.fontSize !== null &&
          props.fontFamily !== null &&
          props.fontFamilyEa !== null &&
          props.fontFamilyCs !== null &&
          props.color !== null
        ) {
          break;
        }

        const defRPr = getDefRPrFromStyle(source, level);
        if (!defRPr) continue;

        if (props.fontSize === null && defRPr.fontSize !== undefined) {
          props.fontSize = defRPr.fontSize;
        }
        if (props.fontFamily === null && defRPr.fontFamily != null) {
          props.fontFamily = resolveThemeFont(defRPr.fontFamily, context.fontScheme);
        }
        if (props.fontFamilyEa === null && defRPr.fontFamilyEa != null) {
          props.fontFamilyEa = resolveThemeFont(defRPr.fontFamilyEa, context.fontScheme);
        }
        if (props.fontFamilyCs === null && defRPr.fontFamilyCs != null) {
          props.fontFamilyCs = resolveThemeFont(defRPr.fontFamilyCs, context.fontScheme);
        }
        if (props.color === null && defRPr.color !== undefined) {
          props.color = defRPr.color;
        }
      }
    }
  }
}

function findMatchingPlaceholderStyle(
  placeholderType: string | undefined,
  placeholderIdx: number | undefined,
  styles: PlaceholderStyleInfo[],
): DefaultTextStyle | undefined {
  if (!placeholderType) return undefined;

  // idx マッチ優先
  if (placeholderIdx !== undefined) {
    const byIdx = styles.find((s) => s.placeholderIdx === placeholderIdx);
    if (byIdx?.lstStyle) return byIdx.lstStyle;
  }

  // type マッチにフォールバック
  const byType = styles.find((s) => s.placeholderType === placeholderType);
  if (byType?.lstStyle) return byType.lstStyle;

  // OOXML 仕様に基づくタイプフォールバック (ctrTitle→title, subTitle→body)
  const fallbackType = getPlaceholderFallbackType(placeholderType);
  if (fallbackType) {
    const byFallback = styles.find((s) => s.placeholderType === fallbackType);
    return byFallback?.lstStyle;
  }

  return undefined;
}

function getPlaceholderFallbackType(type: string): string | undefined {
  switch (type) {
    case "ctrTitle":
      return "title";
    case "subTitle":
      return "body";
    default:
      return undefined;
  }
}

function getTxStyleForPlaceholder(
  placeholderType: string | undefined,
  txStyles?: TxStyles,
): DefaultTextStyle | undefined {
  if (!txStyles) return undefined;
  if (!placeholderType) return txStyles.otherStyle;

  switch (placeholderType) {
    case "title":
    case "ctrTitle":
      return txStyles.titleStyle;
    case "body":
    case "subTitle":
    case "obj":
      return txStyles.bodyStyle;
    default:
      return txStyles.otherStyle;
  }
}

function getDefRPrFromStyle(
  style: DefaultTextStyle | undefined,
  level: number,
): DefaultRunProperties | undefined {
  if (!style) return undefined;
  return style.levels[level]?.defaultRunProperties ?? style.defaultParagraph?.defaultRunProperties;
}
