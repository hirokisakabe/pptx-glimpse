/**
 * Existing-theme color and font scheme editing.
 *
 * The operation updates the typed source immediately and records a field-level patch for the
 * writer. It deliberately does not rebuild the theme part: format-scheme content, relationships,
 * child order, and unsupported XML remain in the preserved raw part.
 */

import type {
  CreatePptxThemeColorSchemeOptions,
  CreatePptxThemeFontSchemeOptions,
  CreatePptxThemeFontSetOptions,
} from "../builder/create-pptx.js";
import { localName, parseXml, type XmlNode } from "../reader/xml.js";
import { sourceHandlesEqual } from "./edit-descriptors.js";
import { requireRawBinaryPart } from "./editing-shared.js";
import type { SourceHandle } from "./handles.js";
import type {
  PptxSourceModel,
  PptxSourceModelUpdateThemeSchemeEdit,
  ThemeColorSlot,
  ThemeFontSetKind,
  ThemeFontSetPatch,
} from "./pptx-source-model.js";
import type { SourceTheme } from "./presentation.js";

/** Partial existing color-scheme update using the #823 six-digit sRGB value contract. */
export type UpdateThemeColorSchemeInput = Omit<CreatePptxThemeColorSchemeOptions, "name">;

/** Partial major/minor font update using the #823 script-specific typeface contract. */
export type UpdateThemeFontSetInput = CreatePptxThemeFontSetOptions;

/** Existing font-scheme fields that can be edited independently. */
export type UpdateThemeFontSchemeInput = Omit<CreatePptxThemeFontSchemeOptions, "name">;

/** A non-empty field-level patch for one existing theme's color and/or font scheme. */
export interface UpdateThemeSchemeInput {
  readonly colorScheme?: UpdateThemeColorSchemeInput;
  readonly fontScheme?: UpdateThemeFontSchemeInput;
}

export const THEME_COLOR_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const satisfies readonly ThemeColorSlot[];

const THEME_COLOR_SLOT_SET: ReadonlySet<string> = new Set(THEME_COLOR_SLOTS);
const FONT_SET_KINDS = ["major", "minor"] as const satisfies readonly ThemeFontSetKind[];
const FONT_FIELDS = ["latin", "eastAsian", "complexScript"] as const;
const COLOR_CHOICE_NAMES: ReadonlySet<string> = new Set([
  "scrgbClr",
  "srgbClr",
  "hslClr",
  "sysClr",
  "schemeClr",
  "prstClr",
]);

const textDecoder = new TextDecoder();

/** Update selected color/font fields on exactly one existing theme handle. */
export function updateThemeScheme(
  source: PptxSourceModel,
  themeHandle: SourceHandle,
  input: UpdateThemeSchemeInput,
): PptxSourceModel {
  const patch = normalizeThemeSchemePatch(input);
  const matches = source.themes
    .map((theme, index) => ({ theme, index }))
    .filter(({ theme }) => sourceHandlesEqual(theme.handle, themeHandle));
  if (matches.length === 0) {
    throw new Error("updateThemeScheme: theme handle was not found in PptxSourceModel source");
  }
  if (matches.length !== 1) {
    throw new Error("updateThemeScheme: theme handle is ambiguous in PptxSourceModel source");
  }
  const match = matches[0];
  if (match === undefined) throw new Error("updateThemeScheme: resolved theme is missing");

  const rawPart = requireRawBinaryPart(source, match.theme.partPath, "updateThemeScheme");
  validateEditableThemeStructure(parseXml(textDecoder.decode(rawPart.bytes)), patch);

  const updatedTheme = applyTypedThemePatch(match.theme, patch);
  const edit = {
    kind: "updateThemeScheme",
    themePartPath: match.theme.partPath,
    ...patch,
  } satisfies PptxSourceModelUpdateThemeSchemeEdit;
  return {
    ...source,
    themes: source.themes.map((theme, index) => (index === match.index ? updatedTheme : theme)),
    edits: [...(source.edits ?? []), edit],
  };
}

function normalizeThemeSchemePatch(
  input: unknown,
): Omit<PptxSourceModelUpdateThemeSchemeEdit, "kind" | "themePartPath"> {
  if (!isRecord(input)) throw new Error("updateThemeScheme: input must be an object");
  assertOnlyKeys(input, new Set(["colorScheme", "fontScheme"]), "input");

  const colorScheme = normalizeColorScheme(input.colorScheme);
  const fontScheme = normalizeFontScheme(input.fontScheme);
  if (colorScheme === undefined && fontScheme === undefined) {
    throw new Error("updateThemeScheme: at least one color or font field must be provided");
  }
  return {
    ...(colorScheme !== undefined ? { colorScheme } : {}),
    ...(fontScheme !== undefined ? { fontScheme } : {}),
  };
}

/** Validate the raw nodes that a normalized patch will touch, without changing the XML tree. */
export function validateEditableThemeStructure(
  root: XmlNode,
  patch: Omit<PptxSourceModelUpdateThemeSchemeEdit, "kind" | "themePartPath">,
  operationName = "updateThemeScheme",
): void {
  const theme = requireUniqueChild(root, "theme", operationName);
  const themeElements = requireUniqueChild(theme, "themeElements", operationName);
  if (patch.colorScheme !== undefined) {
    const colorScheme = requireUniqueChild(themeElements, "clrScheme", operationName);
    for (const slot of Object.keys(patch.colorScheme)) {
      const slotNode = requireUniqueChild(colorScheme, slot, operationName);
      const colorChoices = childEntries(slotNode).filter(([key]) =>
        COLOR_CHOICE_NAMES.has(localName(key)),
      );
      if (colorChoices.length !== 1) {
        throw new Error(
          `${operationName}: theme color slot '${slot}' must contain exactly one supported color`,
        );
      }
    }
  }
  if (patch.fontScheme !== undefined) {
    const fontScheme = requireUniqueChild(themeElements, "fontScheme", operationName);
    for (const [kind, fields] of Object.entries(patch.fontScheme)) {
      if (fields === undefined) continue;
      const fontSet = requireUniqueChild(fontScheme, `${kind}Font`, operationName);
      for (const field of Object.keys(fields)) {
        requireUniqueChild(fontSet, fontElementName(field), operationName);
      }
    }
  }
}

function applyTypedThemePatch(
  theme: SourceTheme,
  patch: Omit<PptxSourceModelUpdateThemeSchemeEdit, "kind" | "themePartPath">,
): SourceTheme {
  let colorScheme = theme.colorScheme;
  if (patch.colorScheme !== undefined) {
    if (colorScheme === undefined) {
      throw new Error("updateThemeScheme: typed theme color scheme is not available");
    }
    colorScheme = {
      ...colorScheme,
      colors: {
        ...colorScheme.colors,
        ...Object.fromEntries(
          Object.entries(patch.colorScheme).map(([slot, hex]) => [
            slot,
            { kind: "srgb" as const, hex },
          ]),
        ),
      },
    };
  }

  let fontScheme = theme.fontScheme;
  if (patch.fontScheme !== undefined) {
    if (fontScheme === undefined) {
      throw new Error("updateThemeScheme: typed theme font scheme is not available");
    }
    const fontPatch: Record<string, string> = {};
    for (const kind of FONT_SET_KINDS) {
      const fields = patch.fontScheme[kind];
      if (fields === undefined) continue;
      for (const field of FONT_FIELDS) {
        const typeface = fields[field];
        if (typeface !== undefined) fontPatch[`${kind}${fontSourceSuffix(field)}`] = typeface;
      }
    }
    fontScheme = { ...fontScheme, ...fontPatch };
  }
  return {
    ...theme,
    ...(colorScheme !== undefined ? { colorScheme } : {}),
    ...(fontScheme !== undefined ? { fontScheme } : {}),
  };
}

function normalizeColorScheme(
  input: unknown,
): Readonly<Partial<Record<ThemeColorSlot, string>>> | undefined {
  if (input === undefined) return undefined;
  if (!isRecord(input)) throw new Error("updateThemeScheme: colorScheme must be an object");
  assertOnlyKeys(input, THEME_COLOR_SLOT_SET, "colorScheme");
  const result: Partial<Record<ThemeColorSlot, string>> = {};
  for (const slot of THEME_COLOR_SLOTS) {
    const value = input[slot];
    if (value === undefined) continue;
    if (typeof value !== "string" || !/^[0-9A-Fa-f]{6}$/.test(value)) {
      throw new Error(`updateThemeScheme: colorScheme.${slot} must be a 6-digit hex color`);
    }
    result[slot] = value.toUpperCase();
  }
  if (Object.keys(result).length === 0) {
    throw new Error("updateThemeScheme: colorScheme must contain at least one color field");
  }
  return result;
}

function normalizeFontScheme(
  input: unknown,
): Readonly<Partial<Record<ThemeFontSetKind, ThemeFontSetPatch>>> | undefined {
  if (input === undefined) return undefined;
  if (!isRecord(input)) throw new Error("updateThemeScheme: fontScheme must be an object");
  assertOnlyKeys(input, new Set(FONT_SET_KINDS), "fontScheme");
  const result: Partial<Record<ThemeFontSetKind, ThemeFontSetPatch>> = {};
  for (const kind of FONT_SET_KINDS) {
    const value = input[kind];
    if (value === undefined) continue;
    if (!isRecord(value))
      throw new Error(`updateThemeScheme: fontScheme.${kind} must be an object`);
    assertOnlyKeys(value, new Set(FONT_FIELDS), `fontScheme.${kind}`);
    const fields: Record<string, string> = {};
    for (const field of FONT_FIELDS) {
      const typeface = value[field];
      if (typeface === undefined) continue;
      if (typeof typeface !== "string") {
        throw new Error(`updateThemeScheme: fontScheme.${kind}.${field} must be a string`);
      }
      const normalized = typeface.trim();
      if (field === "latin" && normalized === "") {
        throw new Error(`updateThemeScheme: fontScheme.${kind}.latin must be a non-empty string`);
      }
      assertValidXmlAttribute(normalized, `fontScheme.${kind}.${field}`);
      fields[field] = normalized;
    }
    if (Object.keys(fields).length === 0) {
      throw new Error(`updateThemeScheme: fontScheme.${kind} must contain at least one font field`);
    }
    result[kind] = fields;
  }
  if (Object.keys(result).length === 0) {
    throw new Error("updateThemeScheme: fontScheme must contain at least one major or minor field");
  }
  return result;
}

function assertValidXmlAttribute(value: string, field: string): void {
  for (let index = 0; index < value.length; index += 1) {
    const codePoint = value.codePointAt(index);
    if (codePoint === undefined) continue;
    if (codePoint > 0xffff) index += 1;
    const valid =
      codePoint >= 0x20 &&
      codePoint <= 0x10ffff &&
      !(codePoint >= 0xd800 && codePoint <= 0xdfff) &&
      codePoint !== 0xfffe &&
      codePoint !== 0xffff;
    if (!valid) {
      throw new Error(
        `updateThemeScheme: ${field} contains a character forbidden in an XML attribute`,
      );
    }
  }
}

function requireUniqueChild(parent: XmlNode, name: string, operationName: string): XmlNode {
  const matches = childEntries(parent).filter(([key]) => localName(key) === name);
  const nodes = matches.flatMap(([, value]) =>
    Array.isArray(value) ? value.filter(isRecord) : isRecord(value) ? [value] : [],
  );
  if (matches.length !== 1 || nodes.length !== 1) {
    throw new Error(`${operationName}: theme must contain exactly one a:${name} element`);
  }
  const node = nodes[0];
  if (node === undefined) throw new Error(`${operationName}: a:${name} element is missing`);
  return node;
}

function childEntries(node: XmlNode): [string, unknown][] {
  return Object.entries(node).filter(([key]) => !key.startsWith("@_") && key !== "#text");
}

function fontElementName(field: string): string {
  if (field === "latin") return "latin";
  if (field === "eastAsian") return "ea";
  if (field === "complexScript") return "cs";
  throw new Error(`updateThemeScheme: unsupported font field '${field}'`);
}

function fontSourceSuffix(field: string): string {
  if (field === "latin") return "Latin";
  if (field === "eastAsian") return "EastAsian";
  if (field === "complexScript") return "ComplexScript";
  throw new Error(`updateThemeScheme: unsupported font field '${field}'`);
}

function assertOnlyKeys(
  value: Record<string, unknown>,
  allowed: ReadonlySet<string>,
  field: string,
): void {
  const unsupported = Object.keys(value).find((key) => !allowed.has(key));
  if (unsupported !== undefined) {
    throw new Error(`updateThemeScheme: unsupported ${field} field '${unsupported}'`);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
