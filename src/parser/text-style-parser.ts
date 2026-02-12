import type {
  DefaultTextStyle,
  DefaultParagraphLevelProperties,
  DefaultRunProperties,
} from "../model/text.js";
import { hundredthPointToPoint } from "../utils/emu.js";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseDefaultRunProperties(defRPr: any): DefaultRunProperties | undefined {
  if (!defRPr) return undefined;

  const result: DefaultRunProperties = {};

  if (defRPr["@_sz"] !== undefined) {
    result.fontSize = hundredthPointToPoint(Number(defRPr["@_sz"]));
  }
  if (defRPr.latin?.["@_typeface"] !== undefined) {
    result.fontFamily = defRPr.latin["@_typeface"];
  }
  if (defRPr.ea?.["@_typeface"] !== undefined) {
    result.fontFamilyEa = defRPr.ea["@_typeface"];
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

  return Object.keys(result).length > 0 ? result : undefined;
}

export function parseParagraphLevelProperties(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  node: any,
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

  const defRPr = parseDefaultRunProperties(node.defRPr);
  if (defRPr) {
    result.defaultRunProperties = defRPr;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * defPPr + lvl1pPr〜lvl9pPr の構造を DefaultTextStyle としてパースする。
 * presentation.xml の defaultTextStyle および slideMaster の titleStyle/bodyStyle/otherStyle で共通利用。
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseListStyle(node: any): DefaultTextStyle | undefined {
  if (!node) return undefined;

  const defaultParagraph = parseParagraphLevelProperties(node.defPPr);

  const levels: (DefaultParagraphLevelProperties | undefined)[] = [];
  for (let i = 1; i <= 9; i++) {
    levels.push(parseParagraphLevelProperties(node[`lvl${i}pPr`]));
  }

  // すべてのレベルが undefined で defaultParagraph もなければ undefined を返す
  if (!defaultParagraph && levels.every((l) => l === undefined)) {
    return undefined;
  }

  return { defaultParagraph, levels };
}
