/**
 * Headless editor commands, selection, warnings, and history.
 *
 * Expected operation rejections use the shared discriminated failure contract exported
 * by this module. Unexpected programmer errors and invariant violations must propagate.
 * The repository-wide ownership, conversion, and catch rules are documented in
 * [the editor error contract](../../../docs/editor-error-contract.md).
 */

import {
  type AddConnectorInput,
  type AddEmptySlideFromLayoutInput,
  type AddTextBoxInput,
  asSourceNodeId,
  type EditableParagraphProperties,
  type EditableParagraphProperty,
  type EditableShapeFill,
  type EditableShapeOutline,
  type EditableTextRunProperties,
  type EditableTextRunProperty,
  type Emu,
  findParagraphBySourceHandle,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  type MoveSlideInput,
  type PptxSourceModel,
  type PptxSourceModelEdit,
  type PptxSourceModelParagraphPropertiesEdit,
  type PptxSourceModelParagraphTextEdit,
  type PptxSourceModelPictureCropEdit,
  type PptxSourceModelShapeFillEdit,
  type PptxSourceModelShapeOutlineEdit,
  type PptxSourceModelShapeTransformEdit,
  type PptxSourceModelTextRunEdit,
  type PptxSourceModelTextRunPropertiesEdit,
  type SetPictureCropInput,
  type SourceChart,
  type SourceHandle,
  type SourceImage,
  type SourceParagraph,
  type SourceShapeNode,
  type SourceSlide,
  type SourceTextRun,
  type SourceTransform,
  type UpdateBubbleChartDataInput,
  type UpdateChartDataInput,
  type UpdateScatterChartDataInput,
  type UpdateThemeSchemeInput,
} from "@pptx-glimpse/document";

import type {
  EditorApplyCommandResult,
  EditorCommandWarning,
  EditorOperationFailure,
} from "./command-contract.js";
import { invalidCommandFailure } from "./command-contract.js";
import {
  applyCommandToDocument,
  type EditorCommand as CommandEditorCommand,
  type GroupableSourceShape,
} from "./commands/index.js";

export type {
  EditorApplyCommandResult,
  EditorCommandWarning,
  EditorOperationErrorCode,
  EditorOperationFailure,
} from "./command-contract.js";
export type {
  AddConnectorCommand,
  AddEmptySlideFromLayoutCommand,
  AddTextBoxCommand,
  ClearParagraphPropertiesCommand,
  ClearPictureCropCommand,
  ClearTextRunPropertiesCommand,
  DeleteShapeCommand,
  DeleteSlideCommand,
  DuplicateSlideCommand,
  GroupableSourceShape,
  GroupShapesCommand,
  MoveShapeCommand,
  MoveShapesAcrossSlidesCommand,
  MoveShapesCommand,
  MoveSlideCommand,
  ReplaceImageCommand,
  ReplaceParagraphPlainTextCommand,
  ReplaceTextRunPlainTextCommand,
  ResizeShapeCommand,
  SetParagraphPropertiesCommand,
  SetPictureCropCommand,
  SetShapeFillCommand,
  SetShapeOutlineCommand,
  SetShapeTransformCommand,
  SetTextRunPropertiesCommand,
  UngroupShapeCommand,
  UpdateBubbleChartDataCommand,
  UpdateChartDataCommand,
  UpdateScatterChartDataCommand,
  UpdateThemeSchemeCommand,
} from "./commands/index.js";

/**
 * All commands accepted by the high-level `apply` and `applyAll` APIs.
 *
 * Each union member is discriminated by `kind`. TypeDoc expands the command interfaces here so
 * consumers of the `pptx-glimpse` re-export can inspect every payload without importing the
 * lower-level editor package.
 *
 * @inlineType ReplaceTextRunPlainTextCommand
 * @inlineType ReplaceParagraphPlainTextCommand
 * @inlineType SetTextRunPropertiesCommand
 * @inlineType ClearTextRunPropertiesCommand
 * @inlineType SetParagraphPropertiesCommand
 * @inlineType ClearParagraphPropertiesCommand
 * @inlineType MoveShapeCommand
 * @inlineType ResizeShapeCommand
 * @inlineType SetShapeTransformCommand
 * @inlineType SetShapeFillCommand
 * @inlineType SetShapeOutlineCommand
 * @inlineType AddTextBoxCommand
 * @inlineType AddConnectorCommand
 * @inlineType DeleteShapeCommand
 * @inlineType GroupShapesCommand
 * @inlineType MoveShapesCommand
 * @inlineType UngroupShapeCommand
 * @inlineType ReplaceImageCommand
 * @inlineType SetPictureCropCommand
 * @inlineType ClearPictureCropCommand
 * @inlineType UpdateChartDataCommand
 * @inlineType UpdateScatterChartDataCommand
 * @inlineType UpdateBubbleChartDataCommand
 * @inlineType AddEmptySlideFromLayoutCommand
 * @inlineType DuplicateSlideCommand
 * @inlineType MoveSlideCommand
 * @inlineType DeleteSlideCommand
 * @inlineType UpdateThemeSchemeCommand
 * @inlineType EditableTextRunProperties
 * @inlineType EditableTextRunProperty
 * @inlineType EditableParagraphProperties
 * @inlineType EditableParagraphProperty
 * @inlineType EditableShapeFill
 * @inlineType EditableShapeOutline
 * @inlineType AddTextBoxBodyPropertiesInput
 * @inlineType AddTextBoxParagraphInput
 * @inlineType AddConnectorConnectionEndpointInput
 * @inlineType AddConnectorOutlineInput
 */
export type EditorCommand = CommandEditorCommand;

export type EditorHistoryResult =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
    }
  | EditorOperationFailure<"empty-undo-stack" | "empty-redo-stack">;

export interface EditorSelection {
  readonly shapeHandle: SourceHandle;
}

export type EditorSelectShapeResult =
  | {
      readonly ok: true;
      readonly selection: EditorSelection;
    }
  | EditorOperationFailure<"invalid-selection">;

interface HistoryEntry {
  readonly before: PptxSourceModel;
  readonly after: PptxSourceModel;
  readonly selectionTransition?: {
    readonly before?: EditorSelection;
    readonly after?: EditorSelection;
  };
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

  replaceTextRunPlainText(run: SourceTextRun, text: string): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "replaceTextRunPlainText",
      run,
      isSourceTextRun,
      findTextRunBySourceHandle,
      (handle) => ({ kind: "replaceTextRunPlainText", handle, text }),
    );
  }

  replaceParagraphPlainText(paragraph: SourceParagraph, text: string): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "replaceParagraphPlainText",
      paragraph,
      isSourceParagraph,
      findParagraphBySourceHandle,
      (handle) => ({ kind: "replaceParagraphPlainText", handle, text }),
    );
  }

  setTextRunProperties(
    run: SourceTextRun,
    properties: EditableTextRunProperties,
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "setTextRunProperties",
      run,
      isSourceTextRun,
      findTextRunBySourceHandle,
      (handle) => ({ kind: "setTextRunProperties", handle, properties }),
    );
  }

  clearTextRunProperties(
    run: SourceTextRun,
    properties: readonly EditableTextRunProperty[],
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "clearTextRunProperties",
      run,
      isSourceTextRun,
      findTextRunBySourceHandle,
      (handle) => ({ kind: "clearTextRunProperties", handle, properties }),
    );
  }

  setParagraphProperties(
    paragraph: SourceParagraph,
    properties: EditableParagraphProperties,
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "setParagraphProperties",
      paragraph,
      isSourceParagraph,
      findParagraphBySourceHandle,
      (handle) => ({ kind: "setParagraphProperties", handle, properties }),
    );
  }

  clearParagraphProperties(
    paragraph: SourceParagraph,
    properties: readonly EditableParagraphProperty[],
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "clearParagraphProperties",
      paragraph,
      isSourceParagraph,
      findParagraphBySourceHandle,
      (handle) => ({ kind: "clearParagraphProperties", handle, properties }),
    );
  }

  moveShape(shape: SourceShapeNode, offsetX: Emu, offsetY: Emu): EditorApplyCommandResult {
    return this.applyToShapeNode("moveShape", shape, (handle) => ({
      kind: "moveShape",
      handle,
      offsetX,
      offsetY,
    }));
  }

  resizeShape(shape: SourceShapeNode, width: Emu, height: Emu): EditorApplyCommandResult {
    return this.applyToShapeNode("resizeShape", shape, (handle) => ({
      kind: "resizeShape",
      handle,
      width,
      height,
    }));
  }

  setShapeTransform(
    shape: SourceShapeNode,
    transform: Pick<SourceTransform, "offsetX" | "offsetY" | "width" | "height">,
  ): EditorApplyCommandResult {
    return this.applyToShapeNode("setShapeTransform", shape, (handle) => ({
      kind: "setShapeTransform",
      handle,
      ...transform,
    }));
  }

  setShapeFill(shape: SourceShapeNode, fill: EditableShapeFill): EditorApplyCommandResult {
    return this.applyToShapeNode("setShapeFill", shape, (handle) => ({
      kind: "setShapeFill",
      handle,
      fill,
    }));
  }

  setShapeOutline(shape: SourceShapeNode, outline: EditableShapeOutline): EditorApplyCommandResult {
    return this.applyToShapeNode("setShapeOutline", shape, (handle) => ({
      kind: "setShapeOutline",
      handle,
      outline,
    }));
  }

  addTextBox(slide: SourceSlide, input: AddTextBoxInput): EditorApplyCommandResult {
    return this.applyToSlide("addTextBox", slide, (slideHandle) => ({
      kind: "addTextBox",
      slideHandle,
      ...input,
    }));
  }

  addConnector(slide: SourceSlide, input: AddConnectorInput): EditorApplyCommandResult {
    return this.applyToSlide("addConnector", slide, (slideHandle) => ({
      kind: "addConnector",
      slideHandle,
      ...input,
    }));
  }

  /** Delete one supported root or nested drawing through its current or stale source node. */
  deleteShape(shape: SourceShapeNode): EditorApplyCommandResult {
    return this.applyToShapeNode("deleteShape", shape, (handle) => ({
      kind: "deleteShape",
      handle,
    }));
  }

  groupShapes(shapes: readonly GroupableSourceShape[]): EditorApplyCommandResult {
    const handles: SourceHandle[] = [];
    for (const shape of shapes) {
      if (!isGroupableSourceShape(shape) || shape.handle === undefined) {
        return invalidSourceNodeFailure("groupShapes", "every source shape requires a handle");
      }
      const current = findShapeNodeBySourceHandle(this.#document, shape.handle);
      if (current === undefined) {
        return invalidSourceNodeFailure(
          "groupShapes",
          "source node handle was not found in the current EditorSession document",
        );
      }
      if (!isGroupableSourceShape(current)) {
        return invalidSourceNodeFailure("groupShapes", "source node kind is not groupable");
      }
      handles.push(shape.handle);
    }
    return this.apply({ kind: "groupShapes", shapeHandles: handles });
  }

  moveShapes(
    shapes: readonly GroupableSourceShape[],
    destinationHandle: SourceHandle,
    beforeShapeHandle?: SourceHandle,
  ): EditorApplyCommandResult {
    const handles: SourceHandle[] = [];
    for (const shape of shapes) {
      if (!isGroupableSourceShape(shape) || shape.handle === undefined) {
        return invalidSourceNodeFailure("moveShapes", "every source shape requires a handle");
      }
      handles.push(shape.handle);
    }
    return this.apply({
      kind: "moveShapes",
      shapeHandles: handles,
      destinationHandle,
      ...(beforeShapeHandle !== undefined ? { beforeShapeHandle } : {}),
    });
  }

  moveShapesAcrossSlides(
    shapes: readonly GroupableSourceShape[],
    destinationSlide: SourceSlide,
    beforeShapeHandle?: SourceHandle,
  ): EditorApplyCommandResult {
    if (destinationSlide.handle === undefined) {
      return invalidSourceNodeFailure(
        "moveShapesAcrossSlides",
        "destination slide requires a handle",
      );
    }
    const handles: SourceHandle[] = [];
    for (const shape of shapes) {
      if (!isGroupableSourceShape(shape) || shape.handle === undefined) {
        return invalidSourceNodeFailure(
          "moveShapesAcrossSlides",
          "every source shape requires a handle",
        );
      }
      handles.push(shape.handle);
    }
    return this.apply({
      kind: "moveShapesAcrossSlides",
      shapeHandles: handles,
      destinationSlideHandle: destinationSlide.handle,
      ...(beforeShapeHandle !== undefined ? { beforeShapeHandle } : {}),
    });
  }

  ungroupShape(group: SourceShapeNode): EditorApplyCommandResult {
    if (group.kind !== "group") {
      return invalidSourceNodeFailure("ungroupShape", "source node is not a group shape");
    }
    return this.applyToShapeNode("ungroupShape", group, (groupHandle) => ({
      kind: "ungroupShape",
      groupHandle,
    }));
  }

  replaceImage(image: SourceImage, bytes: Uint8Array): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "replaceImage",
      image,
      isSourceImage,
      (document, handle) => {
        const shape = findShapeNodeBySourceHandle(document, handle);
        return shape?.kind === "image" ? shape : undefined;
      },
      (handle) => ({ kind: "replaceImage", handle, bytes }),
    );
  }

  setPictureCrop(image: SourceImage, crop: SetPictureCropInput): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "setPictureCrop",
      image,
      isSourceImage,
      (document, handle) => {
        const shape = findShapeNodeBySourceHandle(document, handle);
        return shape?.kind === "image" ? shape : undefined;
      },
      (handle) => ({ ...crop, kind: "setPictureCrop", handle }),
    );
  }

  clearPictureCrop(image: SourceImage): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "clearPictureCrop",
      image,
      isSourceImage,
      (document, handle) => {
        const shape = findShapeNodeBySourceHandle(document, handle);
        return shape?.kind === "image" ? shape : undefined;
      },
      (handle) => ({ kind: "clearPictureCrop", handle }),
    );
  }

  updateChartData(chart: SourceChart, input: UpdateChartDataInput): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "updateChartData",
      chart,
      isSourceChart,
      (document, handle) => {
        const shape = findShapeNodeBySourceHandle(document, handle);
        return shape?.kind === "chart" ? shape : undefined;
      },
      (handle) => ({ kind: "updateChartData", handle, ...input }),
    );
  }

  updateScatterChartData(
    chart: SourceChart,
    input: UpdateScatterChartDataInput,
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "updateScatterChartData",
      chart,
      isSourceChart,
      (document, handle) => {
        const shape = findShapeNodeBySourceHandle(document, handle);
        return shape?.kind === "chart" ? shape : undefined;
      },
      (handle) => ({ kind: "updateScatterChartData", handle, ...input }),
    );
  }

  updateBubbleChartData(
    chart: SourceChart,
    input: UpdateBubbleChartDataInput,
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      "updateBubbleChartData",
      chart,
      isSourceChart,
      (document, handle) => {
        const shape = findShapeNodeBySourceHandle(document, handle);
        return shape?.kind === "chart" ? shape : undefined;
      },
      (handle) => ({ kind: "updateBubbleChartData", handle, ...input }),
    );
  }

  addEmptySlideFromLayout(input: AddEmptySlideFromLayoutInput): EditorApplyCommandResult {
    return this.apply({ kind: "addEmptySlideFromLayout", ...input });
  }

  duplicateSlide(slide: SourceSlide): EditorApplyCommandResult {
    return this.applyToSlide("duplicateSlide", slide, (handle) => ({
      kind: "duplicateSlide",
      handle,
    }));
  }

  moveSlide(slide: SourceSlide, input: MoveSlideInput): EditorApplyCommandResult {
    return this.applyToSlide("moveSlide", slide, (handle) => ({
      kind: "moveSlide",
      handle,
      ...input,
    }));
  }

  deleteSlide(slide: SourceSlide): EditorApplyCommandResult {
    return this.applyToSlide("deleteSlide", slide, (handle) => ({
      kind: "deleteSlide",
      handle,
    }));
  }

  /** Update field-level color/font values for one existing theme handle. */
  updateThemeScheme(handle: SourceHandle, input: UpdateThemeSchemeInput): EditorApplyCommandResult {
    return this.apply({
      kind: "updateThemeScheme",
      handle,
      colorScheme: input.colorScheme,
      fontScheme: input.fontScheme,
    });
  }

  apply(command: EditorCommand): EditorApplyCommandResult {
    return this.applyAll([command]);
  }

  applyAll(commands: readonly EditorCommand[]): EditorApplyCommandResult {
    const before = this.#document;
    const beforeSelection = this.#selection;
    if (commands.length === 0) return { ok: true, document: before };
    const crossSlideMoveIndex = commands.findIndex(
      (command) => command.kind === "moveShapesAcrossSlides",
    );
    if (crossSlideMoveIndex >= 0 && crossSlideMoveIndex !== commands.length - 1) {
      return invalidCommandFailure(
        EXPECTED_BATCH_REJECTION_PREFIXES,
        new Error("moveShapesAcrossSlides: command must be last in an atomic batch"),
      );
    }

    let after = before;
    let afterSelection = beforeSelection;
    const warnings: EditorCommandWarning[] = [];
    for (const command of commands) {
      const result = applyCommandToDocument(after, command);
      if (!result.ok) return result;
      afterSelection = selectionAfterCommand(after, result.document, command, afterSelection);
      after = result.document;
      warnings.push(...collectCommandWarnings(after, [command]));
    }
    try {
      assertNoConflictingParagraphAndRunEdits(after.edits);
    } catch (cause) {
      return invalidCommandFailure(EXPECTED_BATCH_REJECTION_PREFIXES, cause);
    }
    after = normalizeEditorEdits(after);
    const dedupedWarnings = dedupeCommandWarnings(warnings);
    if (after === before) {
      return {
        ok: true,
        document: before,
        ...(dedupedWarnings.length > 0 ? { warnings: dedupedWarnings } : {}),
      };
    }
    this.#document = after;
    this.#selection = reconcileSelection(after, afterSelection);
    const selectionChangedByTopologyCommand = commands.some(
      (command) =>
        command.kind === "groupShapes" ||
        command.kind === "ungroupShape" ||
        command.kind === "moveShapesAcrossSlides",
    );
    this.#undoStack.push({
      before,
      after,
      ...(selectionChangedByTopologyCommand
        ? {
            selectionTransition: {
              ...(beforeSelection !== undefined ? { before: beforeSelection } : {}),
              ...(this.#selection !== undefined ? { after: this.#selection } : {}),
            },
          }
        : {}),
    });
    this.#redoStack.length = 0;

    return {
      ok: true,
      document: after,
      ...(dedupedWarnings.length > 0 ? { warnings: dedupedWarnings } : {}),
    };
  }

  undo(): EditorHistoryResult {
    const entry = this.#undoStack.pop();
    if (entry === undefined) {
      return {
        ok: false,
        code: "empty-undo-stack",
        message: "undo: undo history is empty",
      };
    }

    this.#document = entry.before;
    this.#selection =
      entry.selectionTransition === undefined
        ? reconcileSelection(entry.before, this.#selection)
        : entry.selectionTransition.before;
    this.#redoStack.push(entry);

    return { ok: true, document: entry.before };
  }

  redo(): EditorHistoryResult {
    const entry = this.#redoStack.pop();
    if (entry === undefined) {
      return {
        ok: false,
        code: "empty-redo-stack",
        message: "redo: redo history is empty",
      };
    }

    this.#document = entry.after;
    this.#selection =
      entry.selectionTransition === undefined
        ? reconcileSelection(entry.after, this.#selection)
        : entry.selectionTransition.after;
    this.#undoStack.push(entry);

    return { ok: true, document: entry.after };
  }

  private applyToShapeNode(
    operation: string,
    shape: SourceShapeNode,
    createCommand: (handle: SourceHandle) => EditorCommand,
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      operation,
      shape,
      isSourceShapeNode,
      findShapeNodeBySourceHandle,
      createCommand,
    );
  }

  private applyToSlide(
    operation: string,
    slide: SourceSlide,
    createCommand: (handle: SourceHandle) => EditorCommand,
  ): EditorApplyCommandResult {
    return this.applyToSourceNode(
      operation,
      slide,
      isSourceSlide,
      (document, handle) =>
        document.slides.find((candidate) => sourceHandlesEqual(candidate.handle, handle)),
      createCommand,
    );
  }

  private applyToSourceNode<Node extends { readonly handle?: SourceHandle }>(
    operation: string,
    node: Node,
    isExpectedNode: (value: unknown) => value is Node,
    findCurrentNode: (document: PptxSourceModel, handle: SourceHandle) => Node | undefined,
    createCommand: (handle: SourceHandle) => EditorCommand,
  ): EditorApplyCommandResult {
    if (!isExpectedNode(node)) {
      return invalidSourceNodeFailure(operation, "source node has the wrong target type");
    }
    if (node.handle === undefined) {
      return invalidSourceNodeFailure(operation, "source node does not have a handle");
    }
    if (findCurrentNode(this.#document, node.handle) === undefined) {
      return invalidSourceNodeFailure(
        operation,
        "source node handle was not found in the current EditorSession document",
      );
    }
    return this.apply(createCommand(node.handle));
  }
}

const EXPECTED_BATCH_REJECTION_PREFIXES = [
  "moveShapesAcrossSlides:",
  "replaceTextRunPlainText:",
  "updateTextRunProperties:",
] as const;

export function createEditorSession(document: PptxSourceModel): EditorSession {
  return new EditorSession(document);
}

function invalidSourceNodeFailure(
  operation: string,
  reason: string,
): EditorOperationFailure<"invalid-command"> {
  return {
    ok: false,
    code: "invalid-command",
    message: `${operation}: ${reason}`,
  };
}

function isSourceTextRun(value: unknown): value is SourceTextRun {
  return isObject(value) && value.kind === "textRun";
}

function isSourceParagraph(value: unknown): value is SourceParagraph {
  return isObject(value) && Array.isArray(value.runs);
}

const SOURCE_SHAPE_NODE_KINDS: ReadonlySet<string> = new Set([
  "shape",
  "connector",
  "group",
  "image",
  "table",
  "chart",
  "smartArt",
  "raw",
]);

function isSourceShapeNode(value: unknown): value is SourceShapeNode {
  return (
    isObject(value) && typeof value.kind === "string" && SOURCE_SHAPE_NODE_KINDS.has(value.kind)
  );
}

const GROUPABLE_SOURCE_SHAPE_KINDS: ReadonlySet<string> = new Set([
  "shape",
  "connector",
  "group",
  "image",
  "table",
  "chart",
]);

function isGroupableSourceShape(value: unknown): value is GroupableSourceShape {
  return (
    isObject(value) &&
    typeof value.kind === "string" &&
    GROUPABLE_SOURCE_SHAPE_KINDS.has(value.kind)
  );
}

function isSourceImage(value: unknown): value is SourceImage {
  return isObject(value) && value.kind === "image";
}

function isSourceChart(value: unknown): value is SourceChart {
  return isObject(value) && value.kind === "chart";
}

function isSourceSlide(value: unknown): value is SourceSlide {
  return (
    isObject(value) &&
    typeof value.partPath === "string" &&
    typeof value.layoutPartPath === "string" &&
    Array.isArray(value.shapes)
  );
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function sourceHandlesEqual(left: SourceHandle | undefined, right: SourceHandle): boolean {
  return (
    left !== undefined &&
    left.partPath === right.partPath &&
    left.nodeId === right.nodeId &&
    left.relationshipId === right.relationshipId &&
    left.orderingSlot === right.orderingSlot
  );
}

function selectionAfterCommand(
  before: PptxSourceModel,
  after: PptxSourceModel,
  command: EditorCommand,
  current: EditorSelection | undefined,
): EditorSelection | undefined {
  if (command.kind === "groupShapes") {
    const edit = after.edits?.at(-1);
    if (edit?.kind !== "groupShapes") {
      throw new Error("EditorSession: groupShapes command did not produce a group edit");
    }
    const group = findShapeNodeBySourceHandle(after, {
      partPath: edit.targetPartPath,
      nodeId: asSourceNodeId(edit.groupId),
    });
    if (group?.kind !== "group" || group.handle === undefined) {
      throw new Error("EditorSession: groupShapes command did not produce a selectable group");
    }
    return { shapeHandle: group.handle };
  }
  if (command.kind === "ungroupShape") {
    const group = findShapeNodeBySourceHandle(before, command.groupHandle);
    if (group?.kind !== "group") {
      throw new Error("EditorSession: ungroupShape command target disappeared before selection");
    }
    const firstChildHandle = group.children[0]?.handle;
    return firstChildHandle === undefined ? undefined : { shapeHandle: firstChildHandle };
  }
  if (command.kind === "moveShapesAcrossSlides") {
    if (current === undefined) return undefined;
    const edit = after.edits?.at(-1);
    if (edit?.kind !== "moveShapesAcrossSlides") {
      throw new Error("EditorSession: moveShapesAcrossSlides did not produce a move edit");
    }
    if (current.shapeHandle.partPath !== edit.sourcePartPath) {
      return reconcileSelection(after, current);
    }
    const mapping = edit.nodeIdMappings.find((item) => item.before === current.shapeHandle.nodeId);
    if (mapping === undefined) return reconcileSelection(after, current);
    const moved = findShapeNodeBySourceHandle(after, {
      partPath: edit.destinationPartPath,
      nodeId: mapping.after,
    });
    return moved?.handle === undefined ? undefined : { shapeHandle: moved.handle };
  }
  return reconcileSelection(after, current);
}

function reconcileSelection(
  document: PptxSourceModel,
  selection: EditorSelection | undefined,
): EditorSelection | undefined {
  if (selection === undefined) return undefined;
  return findShapeNodeBySourceHandle(document, selection.shapeHandle) === undefined
    ? undefined
    : selection;
}

function collectCommandWarnings(
  document: PptxSourceModel,
  commands: readonly EditorCommand[],
): EditorCommandWarning[] {
  if (!commands.some((command) => command.kind === "replaceImage")) return [];
  const warnings: EditorCommandWarning[] = [];
  const imageEdits = document.edits?.filter((edit) => edit.kind === "replaceImage") ?? [];

  for (const command of commands) {
    if (command.kind !== "replaceImage") continue;
    let edit: (typeof imageEdits)[number] | undefined;
    for (let index = imageEdits.length - 1; index >= 0; index -= 1) {
      const candidate = imageEdits[index];
      if (
        candidate.handle.partPath === command.handle.partPath &&
        candidate.handle.nodeId === command.handle.nodeId &&
        candidate.handle.relationshipId === command.handle.relationshipId &&
        candidate.handle.orderingSlot === command.handle.orderingSlot
      ) {
        edit = candidate;
        break;
      }
    }
    // Shared replacements are isolated by the document layer. Retain the public warning
    // shape for compatibility, but emit it only if an in-place shared edit is ever supplied.
    if (edit === undefined || edit.mode === "copyOnWrite" || edit.sharedReferenceCount <= 1)
      continue;
    warnings.push({
      code: "shared-media-part",
      mediaPartPath: edit.mediaPartPath,
      referenceCount: edit.sharedReferenceCount,
      message:
        `replaceImage: media part '${edit.mediaPartPath}' is referenced by ` +
        `${edit.sharedReferenceCount} pic shapes; all references now use the replacement image.`,
    });
  }

  return warnings;
}

function dedupeCommandWarnings(warnings: readonly EditorCommandWarning[]): EditorCommandWarning[] {
  const deduped = new Map<string, EditorCommandWarning>();
  for (const warning of warnings) {
    const key = `${warning.code}\0${warning.mediaPartPath}`;
    if (!deduped.has(key)) deduped.set(key, warning);
  }
  return [...deduped.values()];
}

function normalizeEditorEdits(document: PptxSourceModel): PptxSourceModel {
  const edits = document.edits;
  if (edits === undefined) return document;

  const seenTextRuns = new Set<string>();
  const seenTextRunProperties = new Map<string, Set<EditableTextRunProperty>>();
  const seenParagraphProperties = new Map<string, Set<EditableParagraphProperty>>();
  const seenParagraphs = new Set<string>();
  const seenShapeTransforms = new Set<string>();
  const seenShapeFills = new Set<string>();
  const seenPictureCrops = new Set<string>();
  const seenShapeOutlineProperties = new Map<string, Set<EditableShapeOutlineProperty>>();
  const seenChartData = new Set<string>();
  const normalizedShapeOutlineEdits = new Map<string, MutableShapeOutlineEdit>();
  const normalizedReversed: PptxSourceModelEdit[] = [];
  let changed = false;

  for (let index = edits.length - 1; index >= 0; index -= 1) {
    const edit = edits[index];
    switch (edit.kind) {
      case "replaceTextRunPlainText": {
        const key = editHandleNodeKey(edit);
        const paragraphKey = textRunParagraphEditKey(edit);
        if (
          (paragraphKey !== undefined && seenParagraphs.has(paragraphKey)) ||
          seenTextRuns.has(key)
        ) {
          changed = true;
          continue;
        }
        seenTextRuns.add(key);
        normalizedReversed.push(edit);
        continue;
      }
      case "updateTextRunProperties": {
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
      case "updateParagraphProperties": {
        const normalized = normalizeParagraphPropertiesEdit(edit, seenParagraphProperties);
        if (normalized === undefined) {
          changed = true;
          continue;
        }
        if (!editorEditsEqual(normalized, edit)) changed = true;
        normalizedReversed.push(normalized);
        continue;
      }
      case "replaceParagraphPlainText": {
        const key = editHandleNodeKey(edit);
        if (seenParagraphs.has(key)) {
          changed = true;
          continue;
        }
        seenParagraphs.add(key);
        normalizedReversed.push(edit);
        continue;
      }
      case "updateShapeTransform": {
        const key = editHandleNodeKey(edit);
        if (seenShapeTransforms.has(key)) {
          changed = true;
          continue;
        }
        seenShapeTransforms.add(key);
        normalizedReversed.push(edit);
        continue;
      }
      case "updateShapeFill": {
        const key = editHandleNodeKey(edit);
        if (seenShapeFills.has(key)) {
          changed = true;
          continue;
        }
        seenShapeFills.add(key);
        normalizedReversed.push(edit);
        continue;
      }
      case "updatePictureCrop": {
        const key = editHandleNodeKey(edit);
        if (seenPictureCrops.has(key)) {
          changed = true;
          continue;
        }
        seenPictureCrops.add(key);
        normalizedReversed.push(edit);
        continue;
      }
      case "updateShapeOutline": {
        const normalized = normalizeShapeOutlineEdit(
          edit,
          seenShapeOutlineProperties,
          normalizedShapeOutlineEdits,
        );
        if (normalized === undefined || normalized.merged) {
          changed = true;
          continue;
        }
        if (!editorEditsEqual(normalized.edit, edit)) changed = true;
        normalizedReversed.push(normalized.edit);
        continue;
      }
      case "updateChartData":
      case "updateScatterChartData":
      case "updateBubbleChartData": {
        const key = sourceHandleKey(edit.handle);
        if (seenChartData.has(key)) {
          changed = true;
          continue;
        }
        seenChartData.add(key);
        normalizedReversed.push(edit);
        continue;
      }
      case "addChart":
      case "addConnector":
      case "addEmptySlideFromLayout":
      case "addPicture":
      case "addShape":
      case "addSlideLayout":
      case "addTable":
      case "addTextBox":
      case "cloneSlideLayout":
      case "deleteShape":
      case "deleteSlide":
      case "duplicateSlide":
      case "groupShapes":
      case "moveShapes":
      case "moveShapesAcrossSlides":
      case "moveSlide":
      case "reorderShapes":
      case "replaceImage":
      case "setBackground":
      case "setSlideBackground":
      case "ungroupShape":
      case "updateTableCellProperties":
      case "updateThemeScheme":
        normalizedReversed.push(edit);
    }
  }

  if (!changed && normalizedReversed.length === edits.length) return document;
  return {
    ...document,
    edits: normalizedReversed.reverse(),
  };
}

function assertNoConflictingParagraphAndRunEdits(
  edits: readonly PptxSourceModelEdit[] | undefined,
): void {
  if (edits === undefined) return;
  const paragraphEditIndexes = new Map<string, number>();
  edits.forEach((edit, index) => {
    if (edit.kind === "replaceParagraphPlainText") {
      paragraphEditIndexes.set(editHandleNodeKey(edit), index);
    }
  });

  for (const [index, edit] of edits.entries()) {
    if (edit.kind !== "replaceTextRunPlainText" && edit.kind !== "updateTextRunProperties") {
      continue;
    }
    const paragraphKey = textRunParagraphEditKey(edit);
    const paragraphIndex =
      paragraphKey === undefined ? undefined : paragraphEditIndexes.get(paragraphKey);
    if (
      paragraphIndex !== undefined &&
      (edit.kind === "updateTextRunProperties" || paragraphIndex < index)
    ) {
      throw new Error(
        `${edit.kind}: run edits cannot be combined with replaceParagraphPlainText for the same paragraph`,
      );
    }
  }
}

function editorEditsEqual(left: PptxSourceModelEdit, right: PptxSourceModelEdit): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function editHandleNodeKey(
  edit:
    | PptxSourceModelTextRunEdit
    | PptxSourceModelParagraphTextEdit
    | PptxSourceModelParagraphPropertiesEdit
    | PptxSourceModelTextRunPropertiesEdit
    | PptxSourceModelShapeTransformEdit
    | PptxSourceModelShapeFillEdit
    | PptxSourceModelPictureCropEdit
    | PptxSourceModelShapeOutlineEdit,
): string {
  return [
    edit.handle.partPath,
    edit.handle.nodeId ?? "",
    edit.handle.relationshipId ?? "",
    edit.handle.orderingSlot ?? "",
  ].join("\u0000");
}

function sourceHandleKey(handle: SourceHandle): string {
  return [
    handle.partPath,
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot ?? "",
  ].join("\u0000");
}

function textRunParagraphEditKey(
  edit: PptxSourceModelTextRunEdit | PptxSourceModelTextRunPropertiesEdit,
): string | undefined {
  const nodeId = String(edit.handle.nodeId ?? "");
  const match = /^(text:(?:shape:.+|shapeSlot:\d+|table:.+:row:\d+:cell:\d+):p:(\d+)):r:\d+$/.exec(
    nodeId,
  );
  const paragraphNodeId = match?.[1];
  const paragraphOrderingSlot = match?.[2] ?? "";
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

function normalizeParagraphPropertiesEdit(
  edit: PptxSourceModelParagraphPropertiesEdit,
  seenParagraphProperties: Map<string, Set<EditableParagraphProperty>>,
): PptxSourceModelParagraphPropertiesEdit | undefined {
  const key = editHandleNodeKey(edit);
  let seenProperties = seenParagraphProperties.get(key);
  if (seenProperties === undefined) {
    seenProperties = new Set();
    seenParagraphProperties.set(key, seenProperties);
  }

  const set: MutableEditableParagraphProperties = {};
  if (edit.set?.align !== undefined && !seenProperties.has("align")) {
    seenProperties.add("align");
    set.align = edit.set.align;
  }
  if (edit.set?.level !== undefined && !seenProperties.has("level")) {
    seenProperties.add("level");
    set.level = edit.set.level;
  }
  if (edit.set?.bullet !== undefined && !seenProperties.has("bullet")) {
    seenProperties.add("bullet");
    set.bullet = edit.set.bullet;
  }

  const clear = (edit.clear ?? []).filter((property) => !seenProperties.has(property));
  for (const property of clear) seenProperties.add(property);

  if (clear.length === 0 && Object.keys(set).length === 0) return undefined;
  const normalized: PptxSourceModelParagraphPropertiesEdit = {
    kind: "updateParagraphProperties",
    handle: edit.handle,
    ...(Object.keys(set).length > 0 ? { set } : {}),
    ...(clear.length > 0 ? { clear } : {}),
  };
  return normalized;
}

type MutableEditableTextRunProperties = {
  -readonly [K in keyof EditableTextRunProperties]?: EditableTextRunProperties[K];
};

type EditableShapeOutlineProperty = keyof EditableShapeOutline;

type MutableShapeOutline = {
  -readonly [K in keyof EditableShapeOutline]?: EditableShapeOutline[K];
};

interface MutableShapeOutlineEdit {
  readonly kind: "updateShapeOutline";
  readonly handle: SourceHandle;
  outline: MutableShapeOutline;
}

type ShapeOutlineNormalizationResult =
  | {
      readonly edit: MutableShapeOutlineEdit;
      readonly merged: false;
    }
  | {
      readonly merged: true;
    };

function normalizeShapeOutlineEdit(
  edit: PptxSourceModelShapeOutlineEdit,
  seenShapeOutlineProperties: Map<string, Set<EditableShapeOutlineProperty>>,
  normalizedShapeOutlineEdits: Map<string, MutableShapeOutlineEdit>,
): ShapeOutlineNormalizationResult | undefined {
  const key = editHandleNodeKey(edit);
  let seenProperties = seenShapeOutlineProperties.get(key);
  if (seenProperties === undefined) {
    seenProperties = new Set();
    seenShapeOutlineProperties.set(key, seenProperties);
  }

  const outline: MutableShapeOutline = {};
  if (edit.outline.width !== undefined && !seenProperties.has("width")) {
    seenProperties.add("width");
    outline.width = edit.outline.width;
  }
  if (edit.outline.fill !== undefined && !seenProperties.has("fill")) {
    seenProperties.add("fill");
    outline.fill = edit.outline.fill;
  }

  if (Object.keys(outline).length === 0) return undefined;

  const existing = normalizedShapeOutlineEdits.get(key);
  if (existing !== undefined) {
    existing.outline = { ...outline, ...existing.outline };
    return { merged: true };
  }

  const normalized: MutableShapeOutlineEdit = {
    kind: "updateShapeOutline",
    handle: edit.handle,
    outline,
  };
  normalizedShapeOutlineEdits.set(key, normalized);
  return { edit: normalized, merged: false };
}

type MutableEditableParagraphProperties = {
  -readonly [K in keyof EditableParagraphProperties]?: EditableParagraphProperties[K];
};
