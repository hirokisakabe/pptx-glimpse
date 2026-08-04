import { localName, type XmlNode } from "../reader/xml.js";
import type { PptxSourceModelUpdateThemeSchemeEdit } from "../source/index.js";
import {
  drawingMlChildren,
  requireUniqueDrawingMlChild,
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
  const theme = requireUniqueDrawingMlChild(root, {}, "theme", "writePptx");
  const themeElements = requireUniqueDrawingMlChild(
    theme.node,
    theme.namespaces,
    "themeElements",
    "writePptx",
  );

  if (edit.colorScheme !== undefined) {
    const colorScheme = requireUniqueDrawingMlChild(
      themeElements.node,
      themeElements.namespaces,
      "clrScheme",
      "writePptx",
    );
    for (const slot of THEME_COLOR_SLOTS) {
      const hex = edit.colorScheme[slot];
      if (hex === undefined) continue;
      const slotNode = requireUniqueDrawingMlChild(
        colorScheme.node,
        colorScheme.namespaces,
        slot,
        "writePptx",
      );
      const colorChoices = drawingMlChildren(slotNode.node, slotNode.namespaces).filter((child) =>
        COLOR_CHOICE_NAMES.has(localName(child.key)),
      );
      const colorChoice = colorChoices[0];
      if (colorChoice === undefined) {
        throw new Error(`writePptx: theme color slot '${slot}' is missing after validation`);
      }
      const entries = Object.entries(slotNode.node);
      const colorIndex = entries.findIndex(
        ([key, value]) => key === colorChoice.key && value === colorChoice.value,
      );
      if (colorIndex < 0) {
        throw new Error(`writePptx: theme color slot '${slot}' is missing after validation`);
      }
      const oldKey = colorChoice.key;
      const colon = oldKey.indexOf(":");
      const srgbKey = colon === -1 ? "srgbClr" : `${oldKey.slice(0, colon)}:srgbClr`;
      const oldValue = colorChoice.value;
      const newValue = { "@_val": hex };
      const ordered = getXmlChildOrder(slotNode.node).map((entry) =>
        entry.key === oldKey && entry.value === oldValue
          ? { key: srgbKey, value: newValue }
          : entry,
      );
      entries[colorIndex] = [srgbKey, newValue];
      replaceNodeEntries(slotNode.node, entries);
      setXmlChildOrder(slotNode.node, ordered);
    }
  }

  if (edit.fontScheme !== undefined) {
    const fontScheme = requireUniqueDrawingMlChild(
      themeElements.node,
      themeElements.namespaces,
      "fontScheme",
      "writePptx",
    );
    for (const [kind, fields] of Object.entries(edit.fontScheme)) {
      if (fields === undefined) continue;
      const fontSet = requireUniqueDrawingMlChild(
        fontScheme.node,
        fontScheme.namespaces,
        `${kind}Font`,
        "writePptx",
      );
      for (const [field, typeface] of Object.entries(fields)) {
        const font = requireUniqueDrawingMlChild(
          fontSet.node,
          fontSet.namespaces,
          fontElementName(field),
          "writePptx",
        );
        font.node["@_typeface"] = typeface;
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
