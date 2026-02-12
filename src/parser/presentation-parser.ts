import type { SlideSize } from "../model/presentation.js";
import type {
  DefaultTextStyle,
  DefaultParagraphLevelProperties,
  DefaultRunProperties,
} from "../model/text.js";
import { parseXml } from "./xml-parser.js";
import { hundredthPointToPoint } from "../utils/emu.js";

export interface PresentationInfo {
  slideSize: SlideSize;
  slideRIds: string[];
  defaultTextStyle?: DefaultTextStyle;
}

const WARN_PREFIX = "[pptx-glimpse]";
const DEFAULT_SLIDE_WIDTH = 9144000;
const DEFAULT_SLIDE_HEIGHT = 5143500;

export function parsePresentation(xml: string): PresentationInfo {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  const pres = parsed.presentation;

  if (!pres) {
    console.warn(`${WARN_PREFIX} Presentation: missing root element "presentation" in XML`);
    return {
      slideSize: { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT },
      slideRIds: [],
    };
  }

  const sldSz = pres.sldSz;
  let slideSize: SlideSize;
  if (!sldSz || sldSz["@_cx"] === undefined || sldSz["@_cy"] === undefined) {
    console.warn(
      `${WARN_PREFIX} Presentation: sldSz missing, using default ${DEFAULT_SLIDE_WIDTH}x${DEFAULT_SLIDE_HEIGHT} EMU`,
    );
    slideSize = { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
  } else {
    slideSize = {
      width: Number(sldSz["@_cx"]),
      height: Number(sldSz["@_cy"]),
    };
  }

  const sldIdLst = pres.sldIdLst?.sldId ?? [];
  const slideRIds: string[] = sldIdLst
    .map((s: Record<string, string>) => s["@_r:id"] ?? s["@_id"])
    .filter((id: string | undefined) => {
      if (id === undefined) {
        console.warn(
          `${WARN_PREFIX} Presentation: undefined slide relationship ID found, skipping`,
        );
        return false;
      }
      return true;
    });

  const defaultTextStyle = parseDefaultTextStyle(pres.defaultTextStyle);

  return { slideSize, slideRIds, defaultTextStyle };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseDefaultRunProperties(defRPr: any): DefaultRunProperties | undefined {
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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseParagraphLevelProperties(node: any): DefaultParagraphLevelProperties | undefined {
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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseDefaultTextStyle(node: any): DefaultTextStyle | undefined {
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
