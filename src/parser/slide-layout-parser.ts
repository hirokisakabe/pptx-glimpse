import type { Background } from "../model/slide.js";
import type { SlideElement } from "../model/shape.js";
import type { PlaceholderStyleInfo } from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import { parseXml, parseXmlOrdered } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseShapeTree, navigateOrdered } from "./slide-parser.js";
import { buildRelsPath, parseRelationships } from "./relationship-parser.js";
import { parseListStyle } from "./text-style-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";
import { debug } from "../warning-logger.js";

export function parseSlideLayoutBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldLayout) {
    debug("slideLayout.missing", `missing root element "sldLayout" in XML`);
    return null;
  }

  const bg = parsed.sldLayout.cSld?.bg;
  if (!bg) return null;

  const bgPr = bg.bgPr;
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldLayout) {
    debug("slideLayout.missing", `missing root element "sldLayout" in XML`);
    return [];
  }

  const spTree = parsed.sldLayout.cSld?.spTree;
  if (!spTree) return [];

  const relsPath = buildRelsPath(layoutPath);
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map();

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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  const attr = parsed.sldLayout?.["@_showMasterSp"];
  return attr !== "0" && attr !== "false";
}

export function parseSlideLayoutPlaceholderStyles(xml: string): PlaceholderStyleInfo[] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  if (!parsed.sldLayout) return [];

  const spTree = parsed.sldLayout.cSld?.spTree;
  if (!spTree) return [];

  const results: PlaceholderStyleInfo[] = [];
  const shapes = spTree.sp ?? [];

  for (const sp of shapes) {
    const ph = sp.nvSpPr?.nvPr?.ph;
    if (!ph) continue;

    const placeholderType: string = ph["@_type"] ?? "body";
    const placeholderIdx = ph["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;
    const lstStyle = parseListStyle(sp.txBody?.lstStyle);

    results.push({
      placeholderType,
      ...(placeholderIdx !== undefined && { placeholderIdx }),
      ...(lstStyle && { lstStyle }),
    });
  }

  return results;
}
