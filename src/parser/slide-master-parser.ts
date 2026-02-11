import type { ColorMap, ColorSchemeKey } from "../model/theme.js";
import type { Background } from "../model/slide.js";
import type { SlideElement } from "../model/shape.js";
import type { PptxArchive } from "./pptx-reader.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseShapeTree } from "./slide-parser.js";
import { buildRelsPath, parseRelationships } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";

const WARN_PREFIX = "[pptx-glimpse]";

const DEFAULT_COLOR_MAP: ColorMap = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

export function parseSlideMasterColorMap(xml: string): ColorMap {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldMaster) {
    console.warn(`${WARN_PREFIX} SlideMaster: missing root element "sldMaster" in XML`);
    return { ...DEFAULT_COLOR_MAP };
  }

  const clrMap = parsed.sldMaster.clrMap;

  if (!clrMap) return { ...DEFAULT_COLOR_MAP };

  const result: Record<string, string> = {};
  for (const key of Object.keys(DEFAULT_COLOR_MAP)) {
    const val = clrMap[`@_${key}`];
    result[key] = val ?? DEFAULT_COLOR_MAP[key as keyof ColorMap];
  }

  return result as unknown as ColorMap;
}

export function parseSlideMasterBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldMaster) {
    console.warn(`${WARN_PREFIX} SlideMaster: missing root element "sldMaster" in XML`);
    return null;
  }

  const bg = parsed.sldMaster.cSld?.bg;
  if (!bg) return null;

  const bgPr = bg.bgPr;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

export function parseSlideMasterElements(
  xml: string,
  masterPath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): SlideElement[] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldMaster) {
    console.warn(`${WARN_PREFIX} SlideMaster: missing root element "sldMaster" in XML`);
    return [];
  }

  const spTree = parsed.sldMaster.cSld?.spTree;
  if (!spTree) return [];

  const relsPath = buildRelsPath(masterPath);
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map();

  return parseShapeTree(spTree, rels, masterPath, archive, colorResolver);
}

export function getDefaultColorMap(): ColorMap {
  return { ...DEFAULT_COLOR_MAP };
}

export function isValidColorSchemeKey(key: string): key is ColorSchemeKey {
  return [
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
  ].includes(key);
}
