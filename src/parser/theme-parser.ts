import type { Theme, ColorScheme, FontScheme, FormatScheme } from "../model/theme.js";
import type { Fill } from "../model/fill.js";
import type { Outline } from "../model/line.js";
import type { EffectList } from "../model/effect.js";
import { parseXml, parseXmlOrdered, type XmlNode, type XmlOrderedNode } from "./xml-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { parseEffectList } from "./effect-parser.js";
import { ColorResolver } from "../color/color-resolver.js";
import { debug } from "../warning-logger.js";

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
      },
    };
  }

  const themeElements = (parsed.theme as XmlNode).themeElements as XmlNode | undefined;
  if (!themeElements) {
    debug("theme.themeElements", "themeElements not found, using defaults");
    return {
      colorScheme: defaultColorScheme(),
      fontScheme: {
        majorFont: "Calibri",
        minorFont: "Calibri",
        majorFontEa: null,
        minorFontEa: null,
      },
    };
  }

  if (!themeElements.clrScheme) {
    debug("theme.colorScheme", "colorScheme not found, using defaults");
  }
  if (!themeElements.fontScheme) {
    debug("theme.fontScheme", "fontScheme not found, using defaults");
  }

  const colorScheme = parseColorScheme(themeElements.clrScheme as XmlNode);
  const fontScheme = parseFontScheme(themeElements.fontScheme as XmlNode);

  const fmtScheme = parseFmtScheme(
    themeElements.fmtScheme as XmlNode | undefined,
    colorScheme,
    xml,
  );

  return { colorScheme, fontScheme, ...(fmtScheme && { fmtScheme }) };
}

function parseColorScheme(clrScheme: XmlNode): ColorScheme {
  if (!clrScheme) return defaultColorScheme();

  return {
    dk1: extractColor(clrScheme.dk1 as XmlNode),
    lt1: extractColor(clrScheme.lt1 as XmlNode),
    dk2: extractColor(clrScheme.dk2 as XmlNode),
    lt2: extractColor(clrScheme.lt2 as XmlNode),
    accent1: extractColor(clrScheme.accent1 as XmlNode),
    accent2: extractColor(clrScheme.accent2 as XmlNode),
    accent3: extractColor(clrScheme.accent3 as XmlNode),
    accent4: extractColor(clrScheme.accent4 as XmlNode),
    accent5: extractColor(clrScheme.accent5 as XmlNode),
    accent6: extractColor(clrScheme.accent6 as XmlNode),
    hlink: extractColor(clrScheme.hlink as XmlNode),
    folHlink: extractColor(clrScheme.folHlink as XmlNode),
  };
}

function extractColor(colorNode: XmlNode): string {
  if (!colorNode) return "#000000";

  const srgbClr = colorNode.srgbClr as XmlNode | undefined;
  if (srgbClr) {
    return `#${srgbClr["@_val"] as string}`;
  }
  const sysClr = colorNode.sysClr as XmlNode | undefined;
  if (sysClr) {
    return `#${(sysClr["@_lastClr"] as string | undefined) ?? "000000"}`;
  }
  return "#000000";
}

function parseFontScheme(fontScheme: XmlNode): FontScheme {
  if (!fontScheme)
    return { majorFont: "Calibri", minorFont: "Calibri", majorFontEa: null, minorFontEa: null };

  const majorFontNode = fontScheme.majorFont as XmlNode | undefined;
  const minorFontNode = fontScheme.minorFont as XmlNode | undefined;
  const majorFont =
    ((majorFontNode?.latin as XmlNode | undefined)?.["@_typeface"] as string | undefined) ??
    "Calibri";
  const minorFont =
    ((minorFontNode?.latin as XmlNode | undefined)?.["@_typeface"] as string | undefined) ??
    "Calibri";
  const majorFontEa =
    ((majorFontNode?.ea as XmlNode | undefined)?.["@_typeface"] as string | undefined) ?? null;
  const minorFontEa =
    ((minorFontNode?.ea as XmlNode | undefined)?.["@_typeface"] as string | undefined) ?? null;

  return { majorFont, minorFont, majorFontEa, minorFontEa };
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
    fmtSchemeNode.fillStyleLst as XmlNode | undefined,
    colorResolver,
  );
  const bgFillStyles = parseFillStyleListOrdered(
    fmtSchemeOrdered,
    "bgFillStyleLst",
    fmtSchemeNode.bgFillStyleLst as XmlNode | undefined,
    colorResolver,
  );
  const lnStyles = parseLineStyleList(
    fmtSchemeNode.lnStyleLst as XmlNode | undefined,
    colorResolver,
  );
  const effectStyles = parseEffectStyleList(
    fmtSchemeNode.effectStyleLst as XmlNode | undefined,
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
    current = entry[key] as XmlOrderedNode[];
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
  const children = listOrdered[listTag] as XmlOrderedNode[] | undefined;
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
    const fillData = items[idx] as XmlNode | undefined;
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
    const outline = parseOutline(ln as XmlNode, colorResolver);
    if (outline) outlines.push(outline);
  }
  return outlines;
}

function parseEffectStyleList(
  effectStyleLstNode: XmlNode | undefined,
  colorResolver: ColorResolver,
): (EffectList | null)[] {
  if (!effectStyleLstNode) return [];
  const items = (effectStyleLstNode.effectStyle as XmlNode[] | undefined) ?? [];
  return items.map((es) => parseEffectList(es.effectLst as XmlNode, colorResolver));
}
