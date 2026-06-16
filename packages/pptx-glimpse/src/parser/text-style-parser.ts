import type {
  AutoNumScheme,
  DefaultParagraphLevelProperties,
  DefaultRunProperties,
  DefaultTextStyle,
} from "pptx-glimpse-renderer";
import type { FontScheme } from "pptx-glimpse-renderer";
import { hundredthPointToPoint } from "pptx-glimpse-renderer";
import { asEmu, asHundredthPt } from "pptx-glimpse-renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import type { XmlNode } from "./xml-parser.js";

export function parseDefaultRunProperties(
  defRPr: XmlNode,
  colorResolver?: ColorResolver,
): DefaultRunProperties | undefined {
  if (!defRPr) return undefined;

  const result: DefaultRunProperties = {};

  if (defRPr["@_sz"] !== undefined) {
    result.fontSize = hundredthPointToPoint(asHundredthPt(Number(defRPr["@_sz"])));
  }
  const latin = defRPr.latin as XmlNode | undefined;
  if (latin?.["@_typeface"] !== undefined) {
    result.fontFamily = latin["@_typeface"] as string;
  }
  const ea = defRPr.ea as XmlNode | undefined;
  if (ea?.["@_typeface"] !== undefined) {
    result.fontFamilyEa = ea["@_typeface"] as string;
  }
  const cs = defRPr.cs as XmlNode | undefined;
  if (cs?.["@_typeface"] !== undefined) {
    result.fontFamilyCs = cs["@_typeface"] as string;
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

const VALID_AUTO_NUM_SCHEMES = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);

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
    result.marginLeft = asEmu(Number(node["@_marL"]));
  }
  if (node["@_indent"] !== undefined) {
    result.indent = asEmu(Number(node["@_indent"]));
  }

  // Bullet の解析
  if (node.buNone !== undefined) {
    result.bullet = { type: "none" };
  } else if (node.buChar) {
    const buChar = node.buChar as XmlNode;
    result.bullet = { type: "char", char: (buChar["@_char"] as string | undefined) ?? "\u2022" };
  } else if (node.buAutoNum) {
    const buAutoNum = node.buAutoNum as XmlNode;
    const scheme = (buAutoNum["@_type"] as string | undefined) ?? "arabicPeriod";
    result.bullet = {
      type: "autoNum",
      scheme: VALID_AUTO_NUM_SCHEMES.has(scheme) ? (scheme as AutoNumScheme) : "arabicPeriod",
      startAt: Number(buAutoNum["@_startAt"] ?? 1),
    };
  }
  if (node.buFont) {
    const buFont = node.buFont as XmlNode;
    const typeface = (buFont["@_typeface"] as string | undefined) ?? null;
    if (typeface) result.bulletFont = typeface;
  }
  if (colorResolver) {
    const buClr = node.buClr as XmlNode | undefined;
    if (buClr) {
      const color = colorResolver.resolve(buClr);
      if (color) result.bulletColor = color;
    }
  }
  const buSzPct = node.buSzPct as XmlNode | undefined;
  if (buSzPct) {
    result.bulletSizePct = Number(buSzPct["@_val"]);
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
    case "+mj-cs":
      return fontScheme.majorFontCs;
    case "+mn-cs":
      return fontScheme.minorFontCs;
    default:
      return typeface;
  }
}
