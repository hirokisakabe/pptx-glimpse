import { getChild, localName, type XmlNode } from "../reader/xml.js";
import type { PptxSourceModelUpdateThemeSchemeEdit } from "../source/index.js";
import {
  THEME_COLOR_SLOTS,
  validateEditableThemeStructure,
} from "../source/theme-scheme-editing.js";
import { replaceNodeEntries } from "./xml-node-utils.js";
import { getXmlChildOrder, setXmlChildOrder } from "./xml-serialization.js";

const COLOR_CHOICE_NAMES: ReadonlySet<string> = new Set([
  "scrgbClr",
  "srgbClr",
  "hslClr",
  "sysClr",
  "schemeClr",
  "prstClr",
]);

export function applyUpdateThemeSchemeEdit(
  root: XmlNode,
  edit: PptxSourceModelUpdateThemeSchemeEdit,
): void {
  validateEditableThemeStructure(root, edit, "writePptx");
  const themeElements = getChild(getChild(root, "theme"), "themeElements");
  if (themeElements === undefined) {
    throw new Error("writePptx: theme elements are missing after validation");
  }

  if (edit.colorScheme !== undefined) {
    const colorScheme = getChild(themeElements, "clrScheme");
    if (colorScheme === undefined) {
      throw new Error("writePptx: color scheme is missing after validation");
    }
    for (const slot of THEME_COLOR_SLOTS) {
      const hex = edit.colorScheme[slot];
      if (hex === undefined) continue;
      const slotNode = getChild(colorScheme, slot);
      if (slotNode === undefined) {
        throw new Error(`writePptx: theme color slot '${slot}' is missing after validation`);
      }
      const entries = Object.entries(slotNode);
      const colorIndex = entries.findIndex(
        ([key]) => !key.startsWith("@_") && COLOR_CHOICE_NAMES.has(localName(key)),
      );
      const oldKey = entries[colorIndex]?.[0];
      if (colorIndex < 0 || oldKey === undefined) {
        throw new Error(`writePptx: theme color slot '${slot}' is missing after validation`);
      }
      const colon = oldKey.indexOf(":");
      const srgbKey = colon === -1 ? "srgbClr" : `${oldKey.slice(0, colon)}:srgbClr`;
      const oldValue = entries[colorIndex]?.[1];
      const newValue = { "@_val": hex };
      const ordered = getXmlChildOrder(slotNode).map((entry) =>
        entry.key === oldKey && entry.value === oldValue
          ? { key: srgbKey, value: newValue }
          : entry,
      );
      entries[colorIndex] = [srgbKey, newValue];
      replaceNodeEntries(slotNode, entries);
      setXmlChildOrder(slotNode, ordered);
    }
  }

  if (edit.fontScheme !== undefined) {
    const fontScheme = getChild(themeElements, "fontScheme");
    if (fontScheme === undefined) {
      throw new Error("writePptx: font scheme is missing after validation");
    }
    for (const [kind, fields] of Object.entries(edit.fontScheme)) {
      if (fields === undefined) continue;
      const fontSet = getChild(fontScheme, `${kind}Font`);
      if (fontSet === undefined) {
        throw new Error(`writePptx: ${kind} font set is missing after validation`);
      }
      for (const [field, typeface] of Object.entries(fields)) {
        const font = getChild(fontSet, fontElementName(field));
        if (font === undefined) {
          throw new Error(`writePptx: ${kind} ${field} font is missing after validation`);
        }
        font["@_typeface"] = typeface;
      }
    }
  }
}

function fontElementName(field: string): string {
  if (field === "latin") return "latin";
  if (field === "eastAsian") return "ea";
  if (field === "complexScript") return "cs";
  throw new Error(`writePptx: unsupported theme font field '${field}'`);
}
