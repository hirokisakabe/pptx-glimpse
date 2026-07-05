import {
  addEmptySlideFromLayout,
  type AddEmptySlideFromLayoutInput,
  addTextBox,
  type AddTextBoxInput,
  clearTextRunProperties,
  deleteShape,
  deleteSlide,
  duplicateSlide,
  type EditableTextRunProperties,
  type EditableTextRunProperty,
  type Emu,
  findShapeNodeBySourceHandle,
  type PptxSourceModel,
  type PptxSourceModelEdit,
  type PptxSourceModelParagraphTextEdit,
  type PptxSourceModelShapeTransformEdit,
  type PptxSourceModelTextRunEdit,
  type PptxSourceModelTextRunPropertiesEdit,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setTextRunProperties,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTransform,
  updateShapeTransform,
} from "@pptx-glimpse/document";

export {
  type PptxTextBodyProseMirrorCommand,
  type PptxTextBodyProseMirrorDocJson,
  type PptxTextBodyProseMirrorParagraphJson,
  type PptxTextBodyProseMirrorRunMarkJson,
  type PptxTextBodyProseMirrorTextJson,
  pptxTextBodySchema,
  proseMirrorDocJsonToEditorCommands,
  proseMirrorDocJsonToTextBody,
  textBodyToProseMirrorDocJson,
} from "./prosemirror-text-body.js";

export interface ReplaceTextRunPlainTextCommand {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

export interface ReplaceParagraphPlainTextCommand {
  readonly kind: "replaceParagraphPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

export interface SetTextRunPropertiesCommand {
  readonly kind: "setTextRunProperties";
  readonly handle: SourceHandle;
  readonly properties: EditableTextRunProperties;
}

export interface ClearTextRunPropertiesCommand {
  readonly kind: "clearTextRunProperties";
  readonly handle: SourceHandle;
  readonly properties: readonly EditableTextRunProperty[];
}

export interface MoveShapeCommand {
  readonly kind: "moveShape";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
}

export interface ResizeShapeCommand {
  readonly kind: "resizeShape";
  readonly handle: SourceHandle;
  readonly width: Emu;
  readonly height: Emu;
}

export interface SetShapeTransformCommand {
  readonly kind: "setShapeTransform";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}

export interface AddTextBoxCommand extends AddTextBoxInput {
  readonly kind: "addTextBox";
  readonly slideHandle: SourceHandle;
}

export interface DeleteShapeCommand {
  readonly kind: "deleteShape";
  readonly handle: SourceHandle;
}

export interface AddEmptySlideFromLayoutCommand extends AddEmptySlideFromLayoutInput {
  readonly kind: "addEmptySlideFromLayout";
}

export interface DuplicateSlideCommand {
  readonly kind: "duplicateSlide";
  readonly handle: SourceHandle;
}

export interface DeleteSlideCommand {
  readonly kind: "deleteSlide";
  readonly handle: SourceHandle;
}

export type EditorCommand =
  | ReplaceTextRunPlainTextCommand
  | ReplaceParagraphPlainTextCommand
  | SetTextRunPropertiesCommand
  | ClearTextRunPropertiesCommand
  | MoveShapeCommand
  | ResizeShapeCommand
  | SetShapeTransformCommand
  | AddTextBoxCommand
  | DeleteShapeCommand
  | AddEmptySlideFromLayoutCommand
  | DuplicateSlideCommand
  | DeleteSlideCommand;

export type EditorApplyCommandResult =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
    }
  | {
      readonly ok: false;
      readonly code: "invalid-command";
      readonly message: string;
      readonly cause?: unknown;
    };

export type EditorHistoryResult =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
    }
  | {
      readonly ok: false;
      readonly reason: "empty-undo-stack" | "empty-redo-stack";
    };

export interface EditorSelection {
  readonly shapeHandle: SourceHandle;
}

export type EditorSelectShapeResult =
  | {
      readonly ok: true;
      readonly selection: EditorSelection;
    }
  | {
      readonly ok: false;
      readonly code: "invalid-selection";
      readonly message: string;
    };

interface HistoryEntry {
  readonly before: PptxSourceModel;
  readonly after: PptxSourceModel;
}

export class EditorSession {
  #document: PptxSourceModel;
  #selection: EditorSelection | undefined;
  readonly #undoStack: HistoryEntry[] = [];
  readonly #redoStack: HistoryEntry[] = [];

  constructor(document: PptxSourceModel) {
    this.#document = document;
  }

  get document(): PptxSourceModel {
    return this.#document;
  }

  get selection(): EditorSelection | undefined {
    return this.#selection;
  }

  get canUndo(): boolean {
    return this.#undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.#redoStack.length > 0;
  }

  get undoDepth(): number {
    return this.#undoStack.length;
  }

  get redoDepth(): number {
    return this.#redoStack.length;
  }

  selectShape(handle: SourceHandle): EditorSelectShapeResult {
    if (findShapeNodeBySourceHandle(this.#document, handle) === undefined) {
      return {
        ok: false,
        code: "invalid-selection",
        message: "selectShape: shape handle was not found in PptxSourceModel source",
      };
    }

    const selection = { shapeHandle: handle };
    this.#selection = selection;
    return { ok: true, selection };
  }

  deselectShape(): void {
    this.#selection = undefined;
  }

  apply(command: EditorCommand): EditorApplyCommandResult {
    return this.applyAll([command]);
  }

  applyAll(commands: readonly EditorCommand[]): EditorApplyCommandResult {
    const before = this.#document;
    if (commands.length === 0) return { ok: true, document: before };

    try {
      const after = normalizeEditorEdits(
        commands.reduce((document, command) => applyCommandToDocument(document, command), before),
      );
      this.#document = after;
      this.reconcileSelectionAfterDocumentChange();
      this.#undoStack.push({ before, after });
      this.#redoStack.length = 0;

      return { ok: true, document: after };
    } catch (error) {
      return {
        ok: false,
        code: "invalid-command",
        message: error instanceof Error ? error.message : "Editor command was rejected.",
        cause: error,
      };
    }
  }

  undo(): EditorHistoryResult {
    const entry = this.#undoStack.pop();
    if (entry === undefined) return { ok: false, reason: "empty-undo-stack" };

    this.#document = entry.before;
    this.reconcileSelectionAfterDocumentChange();
    this.#redoStack.push(entry);

    return { ok: true, document: entry.before };
  }

  redo(): EditorHistoryResult {
    const entry = this.#redoStack.pop();
    if (entry === undefined) return { ok: false, reason: "empty-redo-stack" };

    this.#document = entry.after;
    this.reconcileSelectionAfterDocumentChange();
    this.#undoStack.push(entry);

    return { ok: true, document: entry.after };
  }

  private reconcileSelectionAfterDocumentChange(): void {
    if (this.#selection === undefined) return;
    if (findShapeNodeBySourceHandle(this.#document, this.#selection.shapeHandle) === undefined) {
      this.#selection = undefined;
    }
  }
}

export function createEditorSession(document: PptxSourceModel): EditorSession {
  return new EditorSession(document);
}

function applyCommandToDocument(
  document: PptxSourceModel,
  command: EditorCommand,
): PptxSourceModel {
  switch (command.kind) {
    case "replaceTextRunPlainText":
      return replaceTextRunPlainText(document, command.handle, command.text);
    case "replaceParagraphPlainText":
      return replaceParagraphPlainText(document, command.handle, command.text);
    case "setTextRunProperties":
      return setTextRunPropertiesCommand(document, command);
    case "clearTextRunProperties":
      return clearTextRunPropertiesCommand(document, command);
    case "moveShape":
      return moveShape(document, command);
    case "resizeShape":
      return resizeShape(document, command);
    case "setShapeTransform":
      return setShapeTransform(document, command);
    case "addTextBox":
      return addTextBoxCommand(document, command);
    case "deleteShape":
      return deleteShape(document, command.handle);
    case "addEmptySlideFromLayout":
      return addEmptySlideFromLayout(document, command);
    case "duplicateSlide":
      return duplicateSlide(document, command.handle);
    case "deleteSlide":
      return deleteSlide(document, command.handle);
  }
}

function addTextBoxCommand(document: PptxSourceModel, command: AddTextBoxCommand): PptxSourceModel {
  requireFiniteEmu(command.offsetX, "addTextBox", "offsetX");
  requireFiniteEmu(command.offsetY, "addTextBox", "offsetY");
  requirePositiveFiniteEmu(command.width, "addTextBox", "width");
  requirePositiveFiniteEmu(command.height, "addTextBox", "height");
  if (typeof command.text !== "string") {
    throw new Error("addTextBox: text must be a string");
  }
  if (command.name !== undefined && command.name.trim() === "") {
    throw new Error("addTextBox: name must be a non-empty string when provided");
  }
  return addTextBox(document, command.slideHandle, command);
}

function setTextRunPropertiesCommand(
  document: PptxSourceModel,
  command: SetTextRunPropertiesCommand,
): PptxSourceModel {
  requireNonEmptyPropertySet(command.properties, "setTextRunProperties");
  validateTextRunPropertySet(command.properties, "setTextRunProperties");
  return setTextRunProperties(document, command.handle, command.properties);
}

function clearTextRunPropertiesCommand(
  document: PptxSourceModel,
  command: ClearTextRunPropertiesCommand,
): PptxSourceModel {
  if (command.properties.length === 0) {
    throw new Error("clearTextRunProperties: properties must contain at least one property name");
  }
  for (const property of command.properties) {
    if (!EDITABLE_TEXT_RUN_PROPERTY_SET.has(property)) {
      throw new Error(`clearTextRunProperties: unsupported text run property '${property}'`);
    }
  }
  return clearTextRunProperties(document, command.handle, command.properties);
}

function moveShape(document: PptxSourceModel, command: MoveShapeCommand): PptxSourceModel {
  requireFiniteEmu(command.offsetX, "moveShape", "offsetX");
  requireFiniteEmu(command.offsetY, "moveShape", "offsetY");

  const current = requireEditableShapeTransform(document, command.handle, "moveShape");
  return updateShapeTransform(document, command.handle, {
    offsetX: command.offsetX,
    offsetY: command.offsetY,
    width: current.width,
    height: current.height,
  });
}

function resizeShape(document: PptxSourceModel, command: ResizeShapeCommand): PptxSourceModel {
  requirePositiveFiniteEmu(command.width, "resizeShape", "width");
  requirePositiveFiniteEmu(command.height, "resizeShape", "height");

  const current = requireEditableShapeTransform(document, command.handle, "resizeShape");
  return updateShapeTransform(document, command.handle, {
    offsetX: current.offsetX,
    offsetY: current.offsetY,
    width: command.width,
    height: command.height,
  });
}

function setShapeTransform(
  document: PptxSourceModel,
  command: SetShapeTransformCommand,
): PptxSourceModel {
  requireFiniteEmu(command.offsetX, "setShapeTransform", "offsetX");
  requireFiniteEmu(command.offsetY, "setShapeTransform", "offsetY");
  requirePositiveFiniteEmu(command.width, "setShapeTransform", "width");
  requirePositiveFiniteEmu(command.height, "setShapeTransform", "height");

  requireEditableShapeTransform(document, command.handle, "setShapeTransform");
  return updateShapeTransform(document, command.handle, {
    offsetX: command.offsetX,
    offsetY: command.offsetY,
    width: command.width,
    height: command.height,
  });
}

const EDITABLE_TEXT_RUN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "fontSize",
  "color",
  "typeface",
] as const satisfies readonly EditableTextRunProperty[];
const EDITABLE_TEXT_RUN_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_TEXT_RUN_PROPERTIES);

function requireNonEmptyPropertySet(
  properties: EditableTextRunProperties,
  commandName: "setTextRunProperties",
): void {
  if (Object.values(properties).every((value) => value === undefined)) {
    throw new Error(`${commandName}: properties must contain at least one defined property`);
  }
}

function validateTextRunPropertySet(
  properties: EditableTextRunProperties,
  commandName: "setTextRunProperties",
): void {
  for (const property of Object.keys(properties)) {
    if (!EDITABLE_TEXT_RUN_PROPERTY_SET.has(property)) {
      throw new Error(`${commandName}: unsupported text run property '${property}'`);
    }
  }
  requireBooleanOrUndefined(properties.bold, commandName, "bold");
  requireBooleanOrUndefined(properties.italic, commandName, "italic");
  requireBooleanOrUndefined(properties.underline, commandName, "underline");
  if (
    properties.fontSize !== undefined &&
    (!Number.isFinite(properties.fontSize) || properties.fontSize <= 0)
  ) {
    throw new Error(`${commandName}: fontSize must be a finite positive pt value`);
  }
  if (properties.typeface !== undefined && properties.typeface.trim() === "") {
    throw new Error(`${commandName}: typeface must be a non-empty string`);
  }
  if (properties.color !== undefined) {
    if (properties.color.kind !== "srgb") {
      throw new Error(`${commandName}: only srgb text run color is supported`);
    }
    if (!/^[0-9A-Fa-f]{6}$/.test(properties.color.hex)) {
      throw new Error(`${commandName}: color.hex must be a 6-digit hex value`);
    }
  }
}

function requireBooleanOrUndefined(
  value: boolean | undefined,
  commandName: "setTextRunProperties",
  fieldName: "bold" | "italic" | "underline",
): void {
  if (value !== undefined && typeof value !== "boolean") {
    throw new Error(`${commandName}: ${fieldName} must be a boolean value`);
  }
}

function requireEditableShapeTransform(
  document: PptxSourceModel,
  handle: SourceHandle,
  commandName: "moveShape" | "resizeShape" | "setShapeTransform",
): SourceTransform {
  const shape = findShapeNodeBySourceHandle(document, handle);
  if (shape === undefined) {
    throw new Error(`${commandName}: shape handle was not found in PptxSourceModel source`);
  }
  if (!hasTransform(shape)) {
    throw new Error(`${commandName}: shape handle does not reference a shape with xfrm`);
  }
  return shape.transform;
}

function requireFiniteEmu(
  value: Emu,
  commandName: "moveShape" | "resizeShape" | "setShapeTransform" | "addTextBox",
  fieldName: string,
): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${commandName}: ${fieldName} must be a finite EMU value`);
  }
}

function requirePositiveFiniteEmu(
  value: Emu,
  commandName: "moveShape" | "resizeShape" | "setShapeTransform" | "addTextBox",
  fieldName: string,
): void {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${commandName}: ${fieldName} must be a finite positive EMU value`);
  }
}

function hasTransform(shape: SourceShapeNode): shape is SourceShapeNode & {
  readonly transform: SourceTransform;
} {
  return shape.kind !== "raw" && shape.transform !== undefined;
}

function normalizeEditorEdits(document: PptxSourceModel): PptxSourceModel {
  const edits = document.edits;
  if (edits === undefined) return document;

  const seenTextRuns = new Set<string>();
  const seenTextRunProperties = new Map<string, Set<EditableTextRunProperty>>();
  const seenParagraphs = new Set<string>();
  const seenShapeTransforms = new Set<string>();
  const normalizedReversed: PptxSourceModelEdit[] = [];
  let changed = false;

  for (let index = edits.length - 1; index >= 0; index -= 1) {
    const edit = edits[index];
    if (edit.kind === "replaceTextRunPlainText") {
      const key = editHandleNodeKey(edit);
      const paragraphKey = textRunParagraphEditKey(edit);
      if (paragraphKey !== undefined && seenParagraphs.has(paragraphKey)) {
        changed = true;
        continue;
      }
      if (seenTextRuns.has(key)) {
        changed = true;
        continue;
      }
      seenTextRuns.add(key);
    }
    if (edit.kind === "updateTextRunProperties") {
      const paragraphKey = textRunParagraphEditKey(edit);
      if (paragraphKey !== undefined && seenParagraphs.has(paragraphKey)) {
        changed = true;
        continue;
      }
      const normalized = normalizeTextRunPropertiesEdit(edit, seenTextRunProperties);
      if (normalized === undefined) {
        changed = true;
        continue;
      }
      if (!editorEditsEqual(normalized, edit)) changed = true;
      normalizedReversed.push(normalized);
      continue;
    }
    if (edit.kind === "replaceParagraphPlainText") {
      const key = editHandleNodeKey(edit);
      if (seenParagraphs.has(key)) {
        changed = true;
        continue;
      }
      seenParagraphs.add(key);
    }
    if (edit.kind === "updateShapeTransform") {
      const key = editHandleNodeKey(edit);
      if (seenShapeTransforms.has(key)) {
        changed = true;
        continue;
      }
      seenShapeTransforms.add(key);
    }
    normalizedReversed.push(edit);
  }

  if (!changed && normalizedReversed.length === edits.length) return document;
  return {
    ...document,
    edits: normalizedReversed.reverse(),
  };
}

function editorEditsEqual(left: PptxSourceModelEdit, right: PptxSourceModelEdit): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function editHandleNodeKey(
  edit:
    | PptxSourceModelTextRunEdit
    | PptxSourceModelParagraphTextEdit
    | PptxSourceModelTextRunPropertiesEdit
    | PptxSourceModelShapeTransformEdit,
): string {
  return [
    edit.handle.partPath,
    edit.handle.nodeId ?? "",
    edit.handle.relationshipId ?? "",
    edit.handle.orderingSlot ?? "",
  ].join("\u0000");
}

function textRunParagraphEditKey(
  edit: PptxSourceModelTextRunEdit | PptxSourceModelTextRunPropertiesEdit,
): string | undefined {
  const nodeId = String(edit.handle.nodeId ?? "");
  const byShapeId = /^(text:shape:.+:p:(\d+)):r:\d+$/.exec(nodeId);
  const byShapeSlot = /^(text:shapeSlot:\d+:p:(\d+)):r:\d+$/.exec(nodeId);
  const paragraphNodeId = byShapeId?.[1] ?? byShapeSlot?.[1];
  const paragraphOrderingSlot = byShapeId?.[2] ?? byShapeSlot?.[2] ?? "";
  if (paragraphNodeId === undefined) return undefined;
  return [
    edit.handle.partPath,
    paragraphNodeId,
    edit.handle.relationshipId ?? "",
    paragraphOrderingSlot,
  ].join("\u0000");
}

function normalizeTextRunPropertiesEdit(
  edit: PptxSourceModelTextRunPropertiesEdit,
  seenTextRunProperties: Map<string, Set<EditableTextRunProperty>>,
): PptxSourceModelTextRunPropertiesEdit | undefined {
  const key = editHandleNodeKey(edit);
  let seenProperties = seenTextRunProperties.get(key);
  if (seenProperties === undefined) {
    seenProperties = new Set();
    seenTextRunProperties.set(key, seenProperties);
  }

  const set: MutableEditableTextRunProperties = {};
  if (edit.set?.bold !== undefined && !seenProperties.has("bold")) {
    seenProperties.add("bold");
    set.bold = edit.set.bold;
  }
  if (edit.set?.italic !== undefined && !seenProperties.has("italic")) {
    seenProperties.add("italic");
    set.italic = edit.set.italic;
  }
  if (edit.set?.underline !== undefined && !seenProperties.has("underline")) {
    seenProperties.add("underline");
    set.underline = edit.set.underline;
  }
  if (edit.set?.fontSize !== undefined && !seenProperties.has("fontSize")) {
    seenProperties.add("fontSize");
    set.fontSize = edit.set.fontSize;
  }
  if (edit.set?.color !== undefined && !seenProperties.has("color")) {
    seenProperties.add("color");
    set.color = edit.set.color;
  }
  if (edit.set?.typeface !== undefined && !seenProperties.has("typeface")) {
    seenProperties.add("typeface");
    set.typeface = edit.set.typeface;
  }

  const clear = (edit.clear ?? []).filter((property) => !seenProperties.has(property));
  for (const property of clear) seenProperties.add(property);

  if (clear.length === 0 && Object.keys(set).length === 0) return undefined;
  const normalized: PptxSourceModelTextRunPropertiesEdit = {
    kind: "updateTextRunProperties",
    handle: edit.handle,
    ...(Object.keys(set).length > 0 ? { set } : {}),
    ...(clear.length > 0 ? { clear } : {}),
  };
  return normalized;
}

type MutableEditableTextRunProperties = {
  -readonly [K in keyof EditableTextRunProperties]?: EditableTextRunProperties[K];
};
