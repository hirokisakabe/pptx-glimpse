import type { EffectList } from "@pptx-glimpse/renderer";
import type { Fill } from "@pptx-glimpse/renderer";
import type { Outline } from "@pptx-glimpse/renderer";
import type { ColorScheme, FontScheme, FormatScheme, Theme } from "@pptx-glimpse/renderer";
import { debug } from "@pptx-glimpse/renderer";

import { ColorResolver } from "../color/color-resolver.js";
import { unsafeXmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { parseEffectList } from "./effect-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { parseXml, parseXmlOrdered, type XmlNode, type XmlOrderedNode } from "./xml-parser.js";

export function parseTheme(xml: string): Theme {
  const parsed = parseXml(xml);

  if (!parsed.theme) {
    debug("theme.missing", `missing root element "theme" in XML`);
    return {
      colorScheme: defaultColorScheme(),
      fontScheme: {
        majorFont: "Calibri",
        minorFont: "Calibri",
        majorFontEa: null,
        minorFontEa: null,
        majorFontCs: null,
        minorFontCs: null,
        majorFontJpan: null,
        minorFontJpan: null,
      },
    };
  }

  const themeElements = unsafeXmlBoundaryAssertion<XmlNode | undefined>(
    unsafeXmlBoundaryAssertion<XmlNode>(parsed.theme).themeElements,
  );
  if (!themeElements) {
    debug("theme.themeElements", "themeElements not found, using defaults");
    return {
      colorScheme: defaultColorScheme(),
      fontScheme: {
        majorFont: "Calibri",
        minorFont: "Calibri",
        majorFontEa: null,
        minorFontEa: null,
        majorFontCs: null,
        minorFontCs: null,
        majorFontJpan: null,
        minorFontJpan: null,
      },
    };
  }

  if (!themeElements.clrScheme) {
    debug("theme.colorScheme", "colorScheme not found, using defaults");
  }
  if (!themeElements.fontScheme) {
    debug("theme.fontScheme", "fontScheme not found, using defaults");
  }

  const colorScheme = parseColorScheme(
    unsafeXmlBoundaryAssertion<XmlNode>(themeElements.clrScheme),
  );
  const fontScheme = parseFontScheme(unsafeXmlBoundaryAssertion<XmlNode>(themeElements.fontScheme));

  const fmtScheme = parseFmtScheme(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(themeElements.fmtScheme),
    colorScheme,
    xml,
  );

  return { colorScheme, fontScheme, ...(fmtScheme && { fmtScheme }) };
}

function parseColorScheme(clrScheme: XmlNode): ColorScheme {
  if (!clrScheme) return defaultColorScheme();

  return {
    dk1: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.dk1)),
    lt1: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.lt1)),
    dk2: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.dk2)),
    lt2: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.lt2)),
    accent1: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.accent1)),
    accent2: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.accent2)),
    accent3: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.accent3)),
    accent4: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.accent4)),
    accent5: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.accent5)),
    accent6: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.accent6)),
    hlink: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.hlink)),
    folHlink: extractColor(unsafeXmlBoundaryAssertion<XmlNode>(clrScheme.folHlink)),
  };
}

function extractColor(colorNode: XmlNode): string {
  if (!colorNode) return "#000000";

  const srgbClr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(colorNode.srgbClr);
  if (srgbClr) {
    return `#${unsafeXmlBoundaryAssertion<string>(srgbClr["@_val"])}`;
  }
  const sysClr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(colorNode.sysClr);
  if (sysClr) {
    return `#${unsafeXmlBoundaryAssertion<string | undefined>(sysClr["@_lastClr"]) ?? "000000"}`;
  }
  return "#000000";
}

function parseFontScheme(fontScheme: XmlNode): FontScheme {
  if (!fontScheme)
    return {
      majorFont: "Calibri",
      minorFont: "Calibri",
      majorFontEa: null,
      minorFontEa: null,
      majorFontCs: null,
      minorFontCs: null,
      majorFontJpan: null,
      minorFontJpan: null,
    };

  const majorFontNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(fontScheme.majorFont);
  const minorFontNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(fontScheme.minorFont);
  const majorFont =
    unsafeXmlBoundaryAssertion<string | undefined>(
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(majorFontNode?.latin)?.["@_typeface"],
    ) ?? "Calibri";
  const minorFont =
    unsafeXmlBoundaryAssertion<string | undefined>(
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(minorFontNode?.latin)?.["@_typeface"],
    ) ?? "Calibri";
  const majorFontEa = resolveEaFont(majorFontNode);
  const minorFontEa = resolveEaFont(minorFontNode);
  const majorFontCs =
    unsafeXmlBoundaryAssertion<string | undefined>(
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(majorFontNode?.cs)?.["@_typeface"],
    ) ?? null;
  const minorFontCs =
    unsafeXmlBoundaryAssertion<string | undefined>(
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(minorFontNode?.cs)?.["@_typeface"],
    ) ?? null;
  const majorFontJpan = findScriptFont(majorFontNode, "Jpan");
  const minorFontJpan = findScriptFont(minorFontNode, "Jpan");

  return {
    majorFont,
    minorFont,
    majorFontEa,
    minorFontEa,
    majorFontCs,
    minorFontCs,
    majorFontJpan,
    minorFontJpan,
  };
}

/**
 * ea タグの typeface を取得し、空文字の場合は script="Jpan" のフォントにフォールバックする。
 */
function resolveEaFont(fontNode: XmlNode | undefined): string | null {
  const eaTypeface = unsafeXmlBoundaryAssertion<string | undefined>(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(fontNode?.ea)?.["@_typeface"],
  );
  if (eaTypeface) return eaTypeface;

  return findScriptFont(fontNode, "Jpan");
}

/**
 * <a:font script="..." typeface="..."> からスクリプトベースのフォント名を取得する。
 */
function findScriptFont(fontNode: XmlNode | undefined, script: string): string | null {
  const fontItems = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(fontNode?.font);
  if (!fontItems) return null;
  for (const f of fontItems) {
    if (f["@_script"] === script && f["@_typeface"]) {
      return unsafeXmlBoundaryAssertion<string>(f["@_typeface"]);
    }
  }
  return null;
}

function defaultColorScheme(): ColorScheme {
  return {
    dk1: "#000000",
    lt1: "#FFFFFF",
    dk2: "#44546A",
    lt2: "#E7E6E6",
    accent1: "#4472C4",
    accent2: "#ED7D31",
    accent3: "#A5A5A5",
    accent4: "#FFC000",
    accent5: "#5B9BD5",
    accent6: "#70AD47",
    hlink: "#0563C1",
    folHlink: "#954F72",
  };
}

const defaultColorMap = {
  bg1: "lt1" as const,
  tx1: "dk1" as const,
  bg2: "lt2" as const,
  tx2: "dk2" as const,
  accent1: "accent1" as const,
  accent2: "accent2" as const,
  accent3: "accent3" as const,
  accent4: "accent4" as const,
  accent5: "accent5" as const,
  accent6: "accent6" as const,
  hlink: "hlink" as const,
  folHlink: "folHlink" as const,
};

const FILL_TAGS = new Set(["solidFill", "gradFill", "pattFill", "blipFill", "noFill"]);

function parseFmtScheme(
  fmtSchemeNode: XmlNode | undefined,
  colorScheme: ColorScheme,
  xml: string,
): FormatScheme | undefined {
  if (!fmtSchemeNode) return undefined;

  const colorResolver = new ColorResolver(colorScheme, defaultColorMap);

  // fillStyleLst/bgFillStyleLst contain mixed fill types — use the ordered parser
  // to preserve their document order (fast-xml-parser groups by tag name)
  const ordered = parseXmlOrdered(xml);
  const fmtSchemeOrdered = navigateThemeOrdered(ordered, ["theme", "themeElements", "fmtScheme"]);

  const fillStyles = parseFillStyleListOrdered(
    fmtSchemeOrdered,
    "fillStyleLst",
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(fmtSchemeNode.fillStyleLst),
    colorResolver,
  );
  const bgFillStyles = parseFillStyleListOrdered(
    fmtSchemeOrdered,
    "bgFillStyleLst",
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(fmtSchemeNode.bgFillStyleLst),
    colorResolver,
  );
  const lnStyles = parseLineStyleList(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(fmtSchemeNode.lnStyleLst),
    colorResolver,
  );
  const effectStyles = parseEffectStyleList(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(fmtSchemeNode.effectStyleLst),
    colorResolver,
  );

  if (
    fillStyles.length === 0 &&
    lnStyles.length === 0 &&
    effectStyles.length === 0 &&
    bgFillStyles.length === 0
  ) {
    return undefined;
  }

  return { fillStyles, lnStyles, effectStyles, bgFillStyles };
}

function navigateThemeOrdered(ordered: XmlOrderedNode[], path: string[]): XmlOrderedNode[] | null {
  let current: XmlOrderedNode[] = ordered;
  for (const key of path) {
    const entry = current.find((item: XmlOrderedNode) => key in item);
    if (!entry) return null;
    current = unsafeXmlBoundaryAssertion<XmlOrderedNode[]>(entry[key]);
    if (!Array.isArray(current)) return null;
  }
  return current;
}

function parseFillStyleListOrdered(
  fmtSchemeOrdered: XmlOrderedNode[] | null,
  listTag: string,
  listNode: XmlNode | undefined,
  colorResolver: ColorResolver,
): Fill[] {
  if (!listNode || !fmtSchemeOrdered) return [];

  const listOrdered = fmtSchemeOrdered.find((c: XmlOrderedNode) => listTag in c);
  if (!listOrdered) return [];
  const children = unsafeXmlBoundaryAssertion<XmlOrderedNode[] | undefined>(listOrdered[listTag]);
  if (!Array.isArray(children)) return [];

  const fills: Fill[] = [];
  const tagCounters: Record<string, number> = {};

  for (const child of children) {
    const tag = Object.keys(child).find((k) => k !== ":@" && FILL_TAGS.has(k));
    if (!tag) continue;

    const idx = tagCounters[tag] ?? 0;
    tagCounters[tag] = idx + 1;

    // Wrap the fill in a container node for parseFillFromNode
    const rawItems = listNode[tag];
    const items = Array.isArray(rawItems) ? rawItems : rawItems ? [rawItems] : [];
    const fillData = unsafeXmlBoundaryAssertion<XmlNode | undefined>(items[idx]);
    if (!fillData) continue;

    const wrapper: XmlNode = { [tag]: fillData };
    const fill = parseFillFromNode(wrapper, colorResolver);
    if (fill) fills.push(fill);
  }

  return fills;
}

function parseLineStyleList(
  lnStyleLstNode: XmlNode | undefined,
  colorResolver: ColorResolver,
): Outline[] {
  if (!lnStyleLstNode) return [];
  const lnItems = lnStyleLstNode.ln;
  const lnArr = Array.isArray(lnItems) ? lnItems : lnItems ? [lnItems] : [];
  const outlines: Outline[] = [];
  for (const ln of lnArr) {
    const outline = parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(ln), colorResolver);
    if (outline) outlines.push(outline);
  }
  return outlines;
}

function parseEffectStyleList(
  effectStyleLstNode: XmlNode | undefined,
  colorResolver: ColorResolver,
): (EffectList | null)[] {
  if (!effectStyleLstNode) return [];
  const items =
    unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(effectStyleLstNode.effectStyle) ?? [];
  return items.map((es) =>
    parseEffectList(unsafeXmlBoundaryAssertion<XmlNode>(es.effectLst), colorResolver),
  );
}
