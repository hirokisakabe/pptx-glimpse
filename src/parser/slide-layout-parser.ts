import type { Background } from "../model/slide.js";
import type { SlideElement } from "../model/shape.js";
import type { PlaceholderStyleInfo } from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import { parseXml, parseXmlOrdered, type XmlNode } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseShapeTree, navigateOrdered } from "./slide-parser.js";
import { buildRelsPath, parseRelationships, type Relationship } from "./relationship-parser.js";
import { parseListStyle } from "./text-style-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";
import { debug } from "../warning-logger.js";

export function parseSlideLayoutBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  const parsed = parseXml(xml);

  const sldLayout = parsed.sldLayout as XmlNode | undefined;
  if (!sldLayout) {
    debug("slideLayout.missing", `missing root element "sldLayout" in XML`);
    return null;
  }

  const cSld = sldLayout.cSld as XmlNode | undefined;
  const bg = cSld?.bg as XmlNode | undefined;
  if (!bg) return null;

  const bgPr = bg.bgPr as XmlNode | undefined;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

export function parseSlideLayoutElements(
  xml: string,
  layoutPath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): SlideElement[] {
  const parsed = parseXml(xml);

  const sldLayout = parsed.sldLayout as XmlNode | undefined;
  if (!sldLayout) {
    debug("slideLayout.missing", `missing root element "sldLayout" in XML`);
    return [];
  }

  const cSld = sldLayout.cSld as XmlNode | undefined;
  const spTree = cSld?.spTree as XmlNode | undefined;
  if (!spTree) return [];

  const relsPath = buildRelsPath(layoutPath);
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const orderedParsed = parseXmlOrdered(xml);
  const orderedSpTree = navigateOrdered(orderedParsed, ["sldLayout", "cSld", "spTree"]);

  return parseShapeTree(
    spTree,
    rels,
    layoutPath,
    archive,
    colorResolver,
    undefined,
    undefined,
    fontScheme,
    orderedSpTree,
  );
}

export function parseSlideLayoutShowMasterSp(xml: string): boolean {
  const parsed = parseXml(xml);
  const sldLayout = parsed.sldLayout as XmlNode | undefined;
  const attr = sldLayout?.["@_showMasterSp"];
  return attr !== "0" && attr !== "false";
}

export function parseSlideLayoutPlaceholderStyles(xml: string): PlaceholderStyleInfo[] {
  const parsed = parseXml(xml);

  const sldLayout = parsed.sldLayout as XmlNode | undefined;
  if (!sldLayout) return [];

  const cSld = sldLayout.cSld as XmlNode | undefined;
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
