import type {
  DefaultTextStyle,
  DefaultParagraphLevelProperties,
  DefaultRunProperties,
} from "../model/text.js";
import type { FontScheme } from "../model/theme.js";
import type { ColorResolver } from "../color/color-resolver.js";
import { hundredthPointToPoint } from "../utils/emu.js";
import type { XmlNode } from "./xml-parser.js";

export function parseDefaultRunProperties(
  defRPr: XmlNode,
  colorResolver?: ColorResolver,
): DefaultRunProperties | undefined {
  if (!defRPr) return undefined;

  const result: DefaultRunProperties = {};

  if (defRPr["@_sz"] !== undefined) {
    result.fontSize = hundredthPointToPoint(Number(defRPr["@_sz"]));
  }
  const latin = defRPr.latin as XmlNode | undefined;
  if (latin?.["@_typeface"] !== undefined) {
    result.fontFamily = latin["@_typeface"] as string;
  }
  const ea = defRPr.ea as XmlNode | undefined;
  if (ea?.["@_typeface"] !== undefined) {
    result.fontFamilyEa = ea["@_typeface"] as string;
  }
  if (defRPr["@_b"] !== undefined) {
    result.bold = defRPr["@_b"] === "1" || defRPr["@_b"] === "true";
  }
  if (defRPr["@_i"] !== undefined) {
    result.italic = defRPr["@_i"] === "1" || defRPr["@_i"] === "true";
  }
  if (defRPr["@_u"] !== undefined) {
    result.underline = defRPr["@_u"] !== "none";
  }
  if (defRPr["@_strike"] !== undefined) {
    result.strikethrough = defRPr["@_strike"] !== "noStrike";
  }

  if (colorResolver) {
    const solidFill = defRPr.solidFill as XmlNode | undefined;
    if (solidFill) {
      const color = colorResolver.resolve(solidFill);
      if (color) {
        result.color = color;
      }
    }
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

export function parseParagraphLevelProperties(
  node: XmlNode,
  colorResolver?: ColorResolver,
): DefaultParagraphLevelProperties | undefined {
  if (!node) return undefined;

  const result: DefaultParagraphLevelProperties = {};

  if (node["@_algn"] !== undefined) {
    result.alignment = node["@_algn"] as "l" | "ctr" | "r" | "just";
  }
  if (node["@_marL"] !== undefined) {
    result.marginLeft = Number(node["@_marL"]);
  }
  if (node["@_indent"] !== undefined) {
    result.indent = Number(node["@_indent"]);
  }

  const defRPr = parseDefaultRunProperties(node.defRPr as XmlNode, colorResolver);
  if (defRPr) {
    result.defaultRunProperties = defRPr;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * defPPr + lvl1pPr〜lvl9pPr の構造を DefaultTextStyle としてパースする。
 * presentation.xml の defaultTextStyle および slideMaster の titleStyle/bodyStyle/otherStyle で共通利用。
 */
export function parseListStyle(
  node: XmlNode,
  colorResolver?: ColorResolver,
): DefaultTextStyle | undefined {
  if (!node) return undefined;

  const defaultParagraph = parseParagraphLevelProperties(node.defPPr as XmlNode, colorResolver);

  const levels: (DefaultParagraphLevelProperties | undefined)[] = [];
  for (let i = 1; i <= 9; i++) {
    levels.push(parseParagraphLevelProperties(node[`lvl${i}pPr`] as XmlNode, colorResolver));
  }

  // すべてのレベルが undefined で defaultParagraph もなければ undefined を返す
  if (!defaultParagraph && levels.every((l) => l === undefined)) {
    return undefined;
  }

  return { defaultParagraph, levels };
}

export function resolveThemeFont(
  typeface: string | null,
  fontScheme?: FontScheme | null,
): string | null {
  if (!typeface || !fontScheme) return typeface;
  switch (typeface) {
    case "+mj-lt":
      return fontScheme.majorFont;
    case "+mn-lt":
      return fontScheme.minorFont;
    case "+mj-ea":
      return fontScheme.majorFontEa;
    case "+mn-ea":
      return fontScheme.minorFontEa;
    default:
      return typeface;
  }
}
