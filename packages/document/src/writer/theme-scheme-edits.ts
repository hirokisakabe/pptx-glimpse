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
      const oldKey = colorChoice.key;
      const colon = oldKey.indexOf(":");
      const srgbKey = colon === -1 ? "srgbClr" : `${oldKey.slice(0, colon)}:srgbClr`;
      const oldValue = colorChoice.value;
      const newValue = drawingMlSrgbValue(oldValue, hex);
      replaceXmlChildByIdentity(slotNode.node, oldKey, oldValue, srgbKey, newValue, slot);
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

function drawingMlSrgbValue(oldValue: unknown, hex: string): XmlNode {
  const declarations =
    typeof oldValue === "object" && oldValue !== null && !Array.isArray(oldValue)
      ? Object.entries(oldValue).filter(([key]) => key === "@_xmlns" || key.startsWith("@_xmlns:"))
      : [];
  return Object.fromEntries([...declarations, ["@_val", hex]]);
}

function replaceXmlChildByIdentity(
  parent: XmlNode,
  oldKey: string,
  oldValue: unknown,
  newKey: string,
  newValue: XmlNode,
  slot: string,
): void {
  const entries = Object.entries(parent);
  let replacementCount = 0;
  const nextEntries = entries.flatMap(([key, groupedValue]): [string, unknown][] => {
    if (key !== oldKey) return [[key, groupedValue]];
    if (!Array.isArray(groupedValue)) {
      if (groupedValue !== oldValue) return [[key, groupedValue]];
      replacementCount += 1;
      return oldKey === newKey ? [[key, newValue]] : [];
    }
    const remaining = groupedValue.filter((member: unknown) => {
      if (member !== oldValue) return true;
      replacementCount += 1;
      return false;
    });
    if (oldKey === newKey) {
      return [
        [key, groupedValue.map((member: unknown) => (member === oldValue ? newValue : member))],
      ];
    }
    return remaining.length === 0 ? [] : [[key, remaining]];
  });
  if (replacementCount !== 1) {
    throw new Error(`writePptx: theme color slot '${slot}' is missing after validation`);
  }
  if (oldKey !== newKey) {
    const newKeyIndex = nextEntries.findIndex(([key]) => key === newKey);
    if (newKeyIndex < 0) {
      nextEntries.push([newKey, newValue]);
    } else {
      const existing = nextEntries[newKeyIndex]?.[1];
      const existingValues: unknown[] = Array.isArray(existing)
        ? existing.map((member: unknown) => member)
        : [existing];
      nextEntries[newKeyIndex] = [newKey, [...existingValues, newValue]];
    }
  }

  const ordered = getXmlChildOrder(parent).map((entry) =>
    entry.key === oldKey && entry.value === oldValue ? { key: newKey, value: newValue } : entry,
  );
  replaceNodeEntries(parent, nextEntries);
  setXmlChildOrder(parent, ordered);
}

function fontElementName(field: string): string {
  if (field === "latin") return "latin";
  if (field === "eastAsian") return "ea";
  if (field === "complexScript") return "cs";
  throw new Error(`writePptx: unsupported theme font field '${field}'`);
}
