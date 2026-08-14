import {
  clearParagraphProperties,
  clearTextRunProperties,
  type EditableParagraphProperties,
  type EditableParagraphProperty,
  type EditableTextRunProperties,
  type EditableTextRunProperty,
  type PptxSourceModel,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setTextRunProperties,
  type SourceHandle,
} from "@pptx-glimpse/document";

import { type ApplyCommandAttempt, attemptCommand } from "../command-contract.js";

/** Replace the text of one source run. @inline */
export interface ReplaceTextRunPlainTextCommand {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

/** Replace all text in one source paragraph. @inline */
export interface ReplaceParagraphPlainTextCommand {
  readonly kind: "replaceParagraphPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

/** Set explicitly supplied text-run properties. @inline */
export interface SetTextRunPropertiesCommand {
  readonly kind: "setTextRunProperties";
  readonly handle: SourceHandle;
  readonly properties: EditableTextRunProperties;
}

/** Clear selected text-run properties so inherited values apply. @inline */
export interface ClearTextRunPropertiesCommand {
  readonly kind: "clearTextRunProperties";
  readonly handle: SourceHandle;
  readonly properties: readonly EditableTextRunProperty[];
}

/** Set explicitly supplied paragraph properties. @inline */
export interface SetParagraphPropertiesCommand {
  readonly kind: "setParagraphProperties";
  readonly handle: SourceHandle;
  readonly properties: EditableParagraphProperties;
}

/** Clear selected paragraph properties so inherited values apply. @inline */
export interface ClearParagraphPropertiesCommand {
  readonly kind: "clearParagraphProperties";
  readonly handle: SourceHandle;
  readonly properties: readonly EditableParagraphProperty[];
}

export type TextEditorCommand =
  | ReplaceTextRunPlainTextCommand
  | ReplaceParagraphPlainTextCommand
  | SetTextRunPropertiesCommand
  | ClearTextRunPropertiesCommand
  | SetParagraphPropertiesCommand
  | ClearParagraphPropertiesCommand;

const EXPECTED_REJECTION_PREFIXES = [
  "replaceTextRunPlainText:",
  "replaceParagraphPlainText:",
  "setTextRunProperties:",
  "clearTextRunProperties:",
  "setParagraphProperties:",
  "clearParagraphProperties:",
  "updateTextRunProperties:",
  "updateParagraphProperties:",
] as const;

const EDITABLE_TEXT_RUN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "fontSize",
  "color",
  "typeface",
] as const satisfies readonly EditableTextRunProperty[];
const EDITABLE_TEXT_RUN_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_TEXT_RUN_PROPERTIES);
const EDITABLE_PARAGRAPH_PROPERTIES = [
  "align",
  "level",
  "bullet",
] as const satisfies readonly EditableParagraphProperty[];
const EDITABLE_PARAGRAPH_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_PARAGRAPH_PROPERTIES);
const PARAGRAPH_ALIGN_VALUES = new Set(["left", "center", "right", "justify"]);
const AUTO_NUM_SCHEMES = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);

export function applyTextCommand(
  document: PptxSourceModel,
  command: TextEditorCommand,
): ApplyCommandAttempt {
  return attemptCommand(EXPECTED_REJECTION_PREFIXES, () => executeTextCommand(document, command));
}

function executeTextCommand(
  document: PptxSourceModel,
  command: TextEditorCommand,
): PptxSourceModel {
  switch (command.kind) {
    case "replaceTextRunPlainText":
      if (typeof command.text !== "string") {
        throw new Error("replaceTextRunPlainText: text must be a string");
      }
      return replaceTextRunPlainText(document, command.handle, command.text);
    case "replaceParagraphPlainText":
      if (typeof command.text !== "string") {
        throw new Error("replaceParagraphPlainText: text must be a string");
      }
      return replaceParagraphPlainText(document, command.handle, command.text);
    case "setTextRunProperties":
      requireNonEmptyPropertySet(command.properties);
      validateTextRunPropertySet(command.properties);
      return setTextRunProperties(document, command.handle, command.properties);
    case "clearTextRunProperties":
      if (command.properties.length === 0) {
        throw new Error(
          "clearTextRunProperties: properties must contain at least one property name",
        );
      }
      for (const property of command.properties) {
        if (!EDITABLE_TEXT_RUN_PROPERTY_SET.has(property)) {
          throw new Error(`clearTextRunProperties: unsupported text run property '${property}'`);
        }
      }
      return clearTextRunProperties(document, command.handle, command.properties);
    case "setParagraphProperties":
      requireNonEmptyParagraphPropertySet(command.properties);
      validateParagraphPropertySet(command.properties);
      return setParagraphProperties(document, command.handle, command.properties);
    case "clearParagraphProperties":
      if (command.properties.length === 0) {
        throw new Error(
          "clearParagraphProperties: properties must contain at least one property name",
        );
      }
      for (const property of command.properties) {
        if (!EDITABLE_PARAGRAPH_PROPERTY_SET.has(property)) {
          throw new Error(`clearParagraphProperties: unsupported paragraph property '${property}'`);
        }
      }
      return clearParagraphProperties(document, command.handle, command.properties);
  }
}

function requireNonEmptyPropertySet(properties: EditableTextRunProperties): void {
  if (Object.values(properties).every((value) => value === undefined)) {
    throw new Error("setTextRunProperties: properties must contain at least one defined property");
  }
}

function validateTextRunPropertySet(properties: EditableTextRunProperties): void {
  for (const property of Object.keys(properties)) {
    if (!EDITABLE_TEXT_RUN_PROPERTY_SET.has(property)) {
      throw new Error(`setTextRunProperties: unsupported text run property '${property}'`);
    }
  }
  requireBooleanOrUndefined(properties.bold, "bold");
  requireBooleanOrUndefined(properties.italic, "italic");
  requireBooleanOrUndefined(properties.underline, "underline");
  if (
    properties.fontSize !== undefined &&
    (!Number.isFinite(properties.fontSize) || properties.fontSize <= 0)
  ) {
    throw new Error("setTextRunProperties: fontSize must be a finite positive pt value");
  }
  if (properties.typeface !== undefined && properties.typeface.trim() === "") {
    throw new Error("setTextRunProperties: typeface must be a non-empty string");
  }
  if (properties.color !== undefined) {
    if (properties.color.kind !== "srgb") {
      throw new Error("setTextRunProperties: only srgb text run color is supported");
    }
    if (!/^[0-9A-Fa-f]{6}$/.test(properties.color.hex)) {
      throw new Error("setTextRunProperties: color.hex must be a 6-digit hex value");
    }
  }
}

function requireNonEmptyParagraphPropertySet(properties: EditableParagraphProperties): void {
  if (Object.values(properties).every((value) => value === undefined)) {
    throw new Error(
      "setParagraphProperties: properties must contain at least one defined property",
    );
  }
}

function validateParagraphPropertySet(properties: EditableParagraphProperties): void {
  for (const property of Object.keys(properties)) {
    if (!EDITABLE_PARAGRAPH_PROPERTY_SET.has(property)) {
      throw new Error(`setParagraphProperties: unsupported paragraph property '${property}'`);
    }
  }
  if (properties.align !== undefined && !PARAGRAPH_ALIGN_VALUES.has(properties.align)) {
    throw new Error("setParagraphProperties: align must be left, center, right, or justify");
  }
  if (
    properties.level !== undefined &&
    (!Number.isInteger(properties.level) || properties.level < 0 || properties.level > 8)
  ) {
    throw new Error("setParagraphProperties: level must be an integer from 0 to 8");
  }
  if (properties.bullet !== undefined) validateParagraphBullet(properties.bullet);
}

function validateParagraphBullet(bullet: NonNullable<EditableParagraphProperties["bullet"]>): void {
  if (bullet.type === "none") return;
  if (bullet.type === "char") {
    if (bullet.char.length === 0) {
      throw new Error("setParagraphProperties: bullet.char must be a non-empty string");
    }
    return;
  }
  if (bullet.type === "autoNum") {
    if (!AUTO_NUM_SCHEMES.has(bullet.scheme)) {
      throw new Error("setParagraphProperties: unsupported bullet auto-numbering scheme");
    }
    if (!Number.isInteger(bullet.startAt) || bullet.startAt < 1) {
      throw new Error("setParagraphProperties: bullet.startAt must be a positive integer");
    }
    return;
  }
  throw new Error("setParagraphProperties: unsupported bullet type");
}

function requireBooleanOrUndefined(
  value: boolean | undefined,
  fieldName: "bold" | "italic" | "underline",
): void {
  if (value !== undefined && typeof value !== "boolean") {
    throw new Error(`setTextRunProperties: ${fieldName} must be a boolean value`);
  }
}
