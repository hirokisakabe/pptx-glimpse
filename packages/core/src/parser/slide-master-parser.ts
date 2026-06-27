import type { SlideElement } from "@pptx-glimpse/renderer";
import type { Background } from "@pptx-glimpse/renderer";
import type { PlaceholderStyleInfo, TxStyles } from "@pptx-glimpse/renderer";
import type { ColorMap, FormatScheme } from "@pptx-glimpse/renderer";
import type { FontScheme } from "@pptx-glimpse/renderer";
import { debug } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { PptxArchive } from "./pptx-reader.js";
import { buildRelsPath, parseRelationships, type Relationship } from "./relationship-parser.js";
import { navigateOrdered, parseGeometry, parseShapeTree, parseTransform } from "./slide-parser.js";
import { parseListStyle } from "./text-style-parser.js";
import { parseXml, parseXmlOrdered, type XmlNode } from "./xml-parser.js";

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

  const sldMaster = unsafeTypeAssertion<XmlNode | undefined>(parsed.sldMaster);
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return { ...DEFAULT_COLOR_MAP };
  }

  const clrMap = unsafeTypeAssertion<XmlNode | undefined>(sldMaster.clrMap);

  if (!clrMap) return { ...DEFAULT_COLOR_MAP };

  const result: Record<string, string> = {};
  for (const key of Object.keys(DEFAULT_COLOR_MAP)) {
    const val = unsafeTypeAssertion<string | undefined>(clrMap[`@_${key}`]);
    result[key] = val ?? DEFAULT_COLOR_MAP[unsafeTypeAssertion<keyof ColorMap>(key)];
  }

  return unsafeTypeAssertion<ColorMap>(result);
}

export function parseSlideMasterBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  const parsed = parseXml(xml);

  const sldMaster = unsafeTypeAssertion<XmlNode | undefined>(parsed.sldMaster);
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return null;
  }

  const cSld = unsafeTypeAssertion<XmlNode | undefined>(sldMaster.cSld);
  const bg = unsafeTypeAssertion<XmlNode | undefined>(cSld?.bg);
  if (!bg) return null;

  const bgPr = unsafeTypeAssertion<XmlNode | undefined>(bg.bgPr);
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
  fmtScheme?: FormatScheme,
): SlideElement[] {
  const parsed = parseXml(xml);

  const sldMaster = unsafeTypeAssertion<XmlNode | undefined>(parsed.sldMaster);
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return [];
  }

  const cSld = unsafeTypeAssertion<XmlNode | undefined>(sldMaster.cSld);
  const spTree = unsafeTypeAssertion<XmlNode | undefined>(cSld?.spTree);
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
    fmtScheme,
  );
}

export function parseSlideMasterTxStyles(
  xml: string,
  colorResolver?: ColorResolver,
): TxStyles | undefined {
  const parsed = parseXml(xml);

  const sldMaster = unsafeTypeAssertion<XmlNode | undefined>(parsed.sldMaster);
  if (!sldMaster) {
    debug("slideMaster.missing", `missing root element "sldMaster" in XML`);
    return undefined;
  }

  const txStyles = unsafeTypeAssertion<XmlNode | undefined>(sldMaster.txStyles);
  if (!txStyles) return undefined;

  const titleStyleNode = unsafeTypeAssertion<XmlNode | undefined>(txStyles.titleStyle);
  const bodyStyleNode = unsafeTypeAssertion<XmlNode | undefined>(txStyles.bodyStyle);
  const otherStyleNode = unsafeTypeAssertion<XmlNode | undefined>(txStyles.otherStyle);
  const titleStyle = titleStyleNode ? parseListStyle(titleStyleNode, colorResolver) : undefined;
  const bodyStyle = bodyStyleNode ? parseListStyle(bodyStyleNode, colorResolver) : undefined;
  const otherStyle = otherStyleNode ? parseListStyle(otherStyleNode, colorResolver) : undefined;

  if (!titleStyle && !bodyStyle && !otherStyle) return undefined;

  return { titleStyle, bodyStyle, otherStyle };
}

export function parseSlideMasterPlaceholderStyles(
  xml: string,
  colorResolver?: ColorResolver,
): PlaceholderStyleInfo[] {
  const parsed = parseXml(xml);
  const sldMaster = unsafeTypeAssertion<XmlNode | undefined>(parsed.sldMaster);
  if (!sldMaster) return [];

  const cSld = unsafeTypeAssertion<XmlNode | undefined>(sldMaster.cSld);
  const spTree = unsafeTypeAssertion<XmlNode | undefined>(cSld?.spTree);
  if (!spTree) return [];

  const results: PlaceholderStyleInfo[] = [];
  const shapes = unsafeTypeAssertion<XmlNode[] | undefined>(spTree.sp) ?? [];

  for (const sp of shapes) {
    const nvSpPr = unsafeTypeAssertion<XmlNode | undefined>(sp.nvSpPr);
    const nvPr = unsafeTypeAssertion<XmlNode | undefined>(nvSpPr?.nvPr);
    const ph = unsafeTypeAssertion<XmlNode | undefined>(nvPr?.ph);
    if (!ph) continue;

    const placeholderType: string = unsafeTypeAssertion<string>(ph["@_type"]) ?? "body";
    const placeholderIdx = ph["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;
    const txBody = unsafeTypeAssertion<XmlNode | undefined>(sp.txBody);
    const lstStyleNode = unsafeTypeAssertion<XmlNode | undefined>(txBody?.lstStyle);
    const lstStyle = lstStyleNode ? parseListStyle(lstStyleNode, colorResolver) : undefined;

    const spPr = unsafeTypeAssertion<XmlNode | undefined>(sp.spPr);
    const transform =
      spPr && typeof spPr === "object"
        ? parseTransform(unsafeTypeAssertion<XmlNode | undefined>(spPr.xfrm))
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
