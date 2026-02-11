import type { Theme, ColorScheme, FontScheme } from "../model/theme.js";
import { parseXml } from "./xml-parser.js";

const WARN_PREFIX = "[pptx-glimpse]";

export function parseTheme(xml: string): Theme {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.theme) {
    console.warn(`${WARN_PREFIX} Theme: missing root element "theme" in XML`);
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

  const themeElements = parsed.theme.themeElements;
  if (!themeElements) {
    console.warn(`${WARN_PREFIX} Theme: themeElements not found, using defaults`);
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
    console.warn(`${WARN_PREFIX} Theme: colorScheme not found, using defaults`);
  }
  if (!themeElements.fontScheme) {
    console.warn(`${WARN_PREFIX} Theme: fontScheme not found, using defaults`);
  }

  const colorScheme = parseColorScheme(themeElements.clrScheme);
  const fontScheme = parseFontScheme(themeElements.fontScheme);

  return { colorScheme, fontScheme };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseColorScheme(clrScheme: any): ColorScheme {
  if (!clrScheme) return defaultColorScheme();

  return {
    dk1: extractColor(clrScheme.dk1),
    lt1: extractColor(clrScheme.lt1),
    dk2: extractColor(clrScheme.dk2),
    lt2: extractColor(clrScheme.lt2),
    accent1: extractColor(clrScheme.accent1),
    accent2: extractColor(clrScheme.accent2),
    accent3: extractColor(clrScheme.accent3),
    accent4: extractColor(clrScheme.accent4),
    accent5: extractColor(clrScheme.accent5),
    accent6: extractColor(clrScheme.accent6),
    hlink: extractColor(clrScheme.hlink),
    folHlink: extractColor(clrScheme.folHlink),
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractColor(colorNode: any): string {
  if (!colorNode) return "#000000";

  if (colorNode.srgbClr) {
    return `#${colorNode.srgbClr["@_val"]}`;
  }
  if (colorNode.sysClr) {
    return `#${colorNode.sysClr["@_lastClr"] ?? "000000"}`;
  }
  return "#000000";
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseFontScheme(fontScheme: any): FontScheme {
  if (!fontScheme)
    return { majorFont: "Calibri", minorFont: "Calibri", majorFontEa: null, minorFontEa: null };

  const majorFont = fontScheme.majorFont?.latin?.["@_typeface"] ?? "Calibri";
  const minorFont = fontScheme.minorFont?.latin?.["@_typeface"] ?? "Calibri";
  const majorFontEa = fontScheme.majorFont?.ea?.["@_typeface"] ?? null;
  const minorFontEa = fontScheme.minorFont?.ea?.["@_typeface"] ?? null;

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
