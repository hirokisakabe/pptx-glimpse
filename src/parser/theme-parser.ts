import type { Theme, ColorScheme, FontScheme } from "../model/theme.js";
import { parseXml, type XmlNode } from "./xml-parser.js";
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

  return { colorScheme, fontScheme };
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
