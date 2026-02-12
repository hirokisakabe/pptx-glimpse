import type { ColorMap, ColorSchemeKey } from "../model/theme.js";
import type { Background } from "../model/slide.js";
import type { SlideElement } from "../model/shape.js";
import type { TxStyles, PlaceholderStyleInfo } from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import { parseXml, parseXmlOrdered, type XmlNode } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseShapeTree, navigateOrdered } from "./slide-parser.js";
import { buildRelsPath, parseRelationships, type Relationship } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";
import { parseListStyle } from "./text-style-parser.js";
import { debug } from "../warning-logger.js";

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
  const parsed = parseXml(xml);

  const sldMaster = parsed.sldMaster as XmlNode | undefined;
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return { ...DEFAULT_COLOR_MAP };
  }

  const clrMap = sldMaster.clrMap as XmlNode | undefined;

  if (!clrMap) return { ...DEFAULT_COLOR_MAP };

  const result: Record<string, string> = {};
  for (const key of Object.keys(DEFAULT_COLOR_MAP)) {
    const val = clrMap[`@_${key}`] as string | undefined;
    result[key] = val ?? DEFAULT_COLOR_MAP[key as keyof ColorMap];
  }

  return result as unknown as ColorMap;
}

export function parseSlideMasterBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  const parsed = parseXml(xml);

  const sldMaster = parsed.sldMaster as XmlNode | undefined;
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return null;
  }

  const cSld = sldMaster.cSld as XmlNode | undefined;
  const bg = cSld?.bg as XmlNode | undefined;
  if (!bg) return null;

  const bgPr = bg.bgPr as XmlNode | undefined;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

export function parseSlideMasterElements(
  xml: string,
  masterPath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): SlideElement[] {
  const parsed = parseXml(xml);

  const sldMaster = parsed.sldMaster as XmlNode | undefined;
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return [];
  }

  const cSld = sldMaster.cSld as XmlNode | undefined;
  const spTree = cSld?.spTree as XmlNode | undefined;
  if (!spTree) return [];

  const relsPath = buildRelsPath(masterPath);
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const orderedParsed = parseXmlOrdered(xml);
  const orderedSpTree = navigateOrdered(orderedParsed, ["sldMaster", "cSld", "spTree"]);

  return parseShapeTree(
    spTree,
    rels,
    masterPath,
    archive,
    colorResolver,
    undefined,
    undefined,
    fontScheme,
    orderedSpTree,
  );
}

export function parseSlideMasterTxStyles(xml: string): TxStyles | undefined {
  const parsed = parseXml(xml);

  const sldMaster = parsed.sldMaster as XmlNode | undefined;
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return undefined;
  }

  const txStyles = sldMaster.txStyles as XmlNode | undefined;
  if (!txStyles) return undefined;

  const titleStyleNode = txStyles.titleStyle as XmlNode | undefined;
  const bodyStyleNode = txStyles.bodyStyle as XmlNode | undefined;
  const otherStyleNode = txStyles.otherStyle as XmlNode | undefined;
  const titleStyle = titleStyleNode ? parseListStyle(titleStyleNode) : undefined;
  const bodyStyle = bodyStyleNode ? parseListStyle(bodyStyleNode) : undefined;
  const otherStyle = otherStyleNode ? parseListStyle(otherStyleNode) : undefined;

  if (!titleStyle && !bodyStyle && !otherStyle) return undefined;

  return { titleStyle, bodyStyle, otherStyle };
}

export function parseSlideMasterPlaceholderStyles(xml: string): PlaceholderStyleInfo[] {
  const parsed = parseXml(xml);
  const sldMaster = parsed.sldMaster as XmlNode | undefined;
  if (!sldMaster) return [];

  const cSld = sldMaster.cSld as XmlNode | undefined;
  const spTree = cSld?.spTree as XmlNode | undefined;
  if (!spTree) return [];

  const results: PlaceholderStyleInfo[] = [];
  const shapes = (spTree.sp as XmlNode[] | undefined) ?? [];

  for (const sp of shapes) {
    const nvSpPr = sp.nvSpPr as XmlNode | undefined;
    const nvPr = nvSpPr?.nvPr as XmlNode | undefined;
    const ph = nvPr?.ph as XmlNode | undefined;
    if (!ph) continue;

    const placeholderType: string = (ph["@_type"] as string) ?? "body";
    const placeholderIdx = ph["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;
    const txBody = sp.txBody as XmlNode | undefined;
    const lstStyleNode = txBody?.lstStyle as XmlNode | undefined;
    const lstStyle = lstStyleNode ? parseListStyle(lstStyleNode) : undefined;

    results.push({
      placeholderType,
      ...(placeholderIdx !== undefined && { placeholderIdx }),
      ...(lstStyle && { lstStyle }),
    });
  }

  return results;
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
