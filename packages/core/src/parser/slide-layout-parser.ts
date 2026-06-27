import type { SlideElement } from "@pptx-glimpse/renderer";
import type { Background } from "@pptx-glimpse/renderer";
import type { PlaceholderStyleInfo } from "@pptx-glimpse/renderer";
import type { FontScheme, FormatScheme } from "@pptx-glimpse/renderer";
import { debug } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import { unsafeXmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { PptxArchive } from "./pptx-reader.js";
import { buildRelsPath, parseRelationships, type Relationship } from "./relationship-parser.js";
import { navigateOrdered, parseGeometry, parseShapeTree, parseTransform } from "./slide-parser.js";
import { parseListStyle } from "./text-style-parser.js";
import { parseXml, parseXmlOrdered, type XmlNode } from "./xml-parser.js";

export function parseSlideLayoutBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  const parsed = parseXml(xml);

  const sldLayout = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.sldLayout);
  if (!sldLayout) {
    debug("slideLayout.missing", `missing root element "sldLayout" in XML`);
    return null;
  }

  const cSld = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sldLayout.cSld);
  const bg = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cSld?.bg);
  if (!bg) return null;

  const bgPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(bg.bgPr);
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
  fmtScheme?: FormatScheme,
): SlideElement[] {
  const parsed = parseXml(xml);

  const sldLayout = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.sldLayout);
  if (!sldLayout) {
    debug("slideLayout.missing", `missing root element "sldLayout" in XML`);
    return [];
  }

  const cSld = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sldLayout.cSld);
  const spTree = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cSld?.spTree);
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
    fmtScheme,
  );
}

export function parseSlideLayoutShowMasterSp(xml: string): boolean {
  const parsed = parseXml(xml);
  const sldLayout = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.sldLayout);
  const attr = sldLayout?.["@_showMasterSp"];
  return attr !== "0" && attr !== "false";
}

export function parseSlideLayoutPlaceholderStyles(
  xml: string,
  colorResolver?: ColorResolver,
): PlaceholderStyleInfo[] {
  const parsed = parseXml(xml);

  const sldLayout = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.sldLayout);
  if (!sldLayout) return [];

  const cSld = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sldLayout.cSld);
  const spTree = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cSld?.spTree);
  if (!spTree) return [];

  const results: PlaceholderStyleInfo[] = [];
  const shapes = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.sp) ?? [];

  for (const sp of shapes) {
    const nvSpPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.nvSpPr);
    const nvPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvSpPr?.nvPr);
    const ph = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvPr?.ph);
    if (!ph) continue;

    const placeholderType: string = unsafeXmlBoundaryAssertion<string>(ph["@_type"]) ?? "body";
    const placeholderIdx = ph["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;
    const txBody = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.txBody);
    const lstStyleNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(txBody?.lstStyle);
    const lstStyle = lstStyleNode ? parseListStyle(lstStyleNode, colorResolver) : undefined;

    const spPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.spPr);
    const transform =
      spPr && typeof spPr === "object"
        ? parseTransform(unsafeXmlBoundaryAssertion<XmlNode | undefined>(spPr.xfrm))
        : null;
    const geometry = spPr && typeof spPr === "object" ? parseGeometry(spPr) : undefined;

    results.push({
      placeholderType,
      ...(placeholderIdx !== undefined && { placeholderIdx }),
      ...(lstStyle && { lstStyle }),
      ...(transform && { transform }),
      ...(geometry && { geometry }),
    });
  }

  return results;
}
