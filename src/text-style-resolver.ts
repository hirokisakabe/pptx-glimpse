import type {
  DefaultTextStyle,
  DefaultRunProperties,
  TxStyles,
  PlaceholderStyleInfo,
} from "./model/text.js";
import type { SlideElement, ShapeElement } from "./model/shape.js";
import type { FontScheme } from "./model/theme.js";
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

    for (const run of paragraph.runs) {
      const props = run.properties;

      for (const source of chainSources) {
        if (
          props.fontSize !== null &&
          props.fontFamily !== null &&
          props.fontFamilyEa !== null &&
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
  return byType?.lstStyle;
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
