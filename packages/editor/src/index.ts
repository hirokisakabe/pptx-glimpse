/**
 * Headless editor commands, selection, warnings, and history.
 *
 * Expected operation rejections use the shared discriminated failure contract exported
 * by this module. Unexpected programmer errors and invariant violations must propagate.
 * The repository-wide ownership, conversion, and catch rules are documented in
 * [the editor error contract](../../../docs/editor-error-contract.md).
 */

import {
  addConnector,
  type AddConnectorInput,
  addEmptySlideFromLayout,
  type AddEmptySlideFromLayoutInput,
  addTextBox,
  type AddTextBoxInput,
  asSourceNodeId,
  clearParagraphProperties,
  clearPictureCrop,
  clearTextRunProperties,
  deleteShape,
  deleteSlide,
  duplicateSlide,
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
  groupShapes,
  moveSlide,
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
  replaceImageBytes,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setPictureCrop,
  type SetPictureCropInput,
  setShapeFill,
  setShapeOutline,
  setTextRunProperties,
  type SourceChart,
  type SourceConnector,
  type SourceGroup,
  type SourceHandle,
  type SourceImage,
  type SourceParagraph,
  type SourceShape,
  type SourceShapeNode,
  type SourceSlide,
  type SourceTable,
  type SourceTextRun,
  type SourceTransform,
  ungroupShape,
  updateChartData,
  type UpdateChartDataInput,
  updateShapeTransform,
} from "@pptx-glimpse/document";

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

/** Move one shape without resizing it. @inline */
export interface MoveShapeCommand {
  readonly kind: "moveShape";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
}

/** Resize one shape without moving its origin. @inline */
export interface ResizeShapeCommand {
  readonly kind: "resizeShape";
  readonly handle: SourceHandle;
  readonly width: Emu;
  readonly height: Emu;
}

/** Replace a shape's position and size together. @inline */
export interface SetShapeTransformCommand {
  readonly kind: "setShapeTransform";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}

/** Replace a shape's fill properties. @inline */
export interface SetShapeFillCommand {
  readonly kind: "setShapeFill";
  readonly handle: SourceHandle;
  readonly fill: EditableShapeFill;
}

/** Update the specified outline properties while preserving omitted properties. @inline */
export interface SetShapeOutlineCommand {
  readonly kind: "setShapeOutline";
  readonly handle: SourceHandle;
  readonly outline: EditableShapeOutline;
}

/** Add a text box to one slide. @inline */
export interface AddTextBoxCommand extends AddTextBoxInput {
  readonly kind: "addTextBox";
  readonly slideHandle: SourceHandle;
}

/** Add a connector to one slide. @inline */
export interface AddConnectorCommand extends AddConnectorInput {
  readonly kind: "addConnector";
  readonly slideHandle: SourceHandle;
}

/** Delete one shape. @inline */
export interface DeleteShapeCommand {
  readonly kind: "deleteShape";
  readonly handle: SourceHandle;
}

/** Group two or more consecutive sibling drawings into one native DrawingML group. @inline */
export interface GroupShapesCommand {
  readonly kind: "groupShapes";
  readonly shapeHandles: readonly SourceHandle[];
}

/** Expand one losslessly ungroupable native DrawingML group. @inline */
export interface UngroupShapeCommand {
  readonly kind: "ungroupShape";
  readonly groupHandle: SourceHandle;
}

/** Source drawing kinds supported by the typed group convenience method. */
export type GroupableSourceShape =
  | SourceShape
  | SourceConnector
  | SourceImage
  | SourceTable
  | SourceChart
  | SourceGroup;

/** Replace the media bytes referenced by one image shape. @inline */
export interface ReplaceImageCommand {
  readonly kind: "replaceImage";
  readonly handle: SourceHandle;
  readonly bytes: Uint8Array;
}

/** Set crop insets on one stretch-filled picture. @inline */
export interface SetPictureCropCommand extends SetPictureCropInput {
  readonly kind: "setPictureCrop";
  readonly handle: SourceHandle;
}

/** Remove the crop rectangle from one stretch-filled picture. @inline */
export interface ClearPictureCropCommand {
  readonly kind: "clearPictureCrop";
  readonly handle: SourceHandle;
}

/** Update the series data of one chart. @inline */
export interface UpdateChartDataCommand extends UpdateChartDataInput {
  readonly kind: "updateChartData";
  readonly handle: SourceHandle;
}

/** Add an empty slide based on an existing layout. @inline */
export interface AddEmptySlideFromLayoutCommand extends AddEmptySlideFromLayoutInput {
  readonly kind: "addEmptySlideFromLayout";
}

/** Duplicate one slide immediately after its source. @inline */
export interface DuplicateSlideCommand {
  readonly kind: "duplicateSlide";
  readonly handle: SourceHandle;
}

/** Move one slide to a new zero-based array position. @inline */
export interface MoveSlideCommand extends MoveSlideInput {
  readonly kind: "moveSlide";
  readonly handle: SourceHandle;
}

/** Delete one slide. @inline */
export interface DeleteSlideCommand {
  readonly kind: "deleteSlide";
  readonly handle: SourceHandle;
}

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
 * @inlineType UngroupShapeCommand
 * @inlineType ReplaceImageCommand
 * @inlineType SetPictureCropCommand
 * @inlineType ClearPictureCropCommand
 * @inlineType UpdateChartDataCommand
 * @inlineType AddEmptySlideFromLayoutCommand
 * @inlineType DuplicateSlideCommand
 * @inlineType MoveSlideCommand
 * @inlineType DeleteSlideCommand
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
export type EditorCommand =
  | ReplaceTextRunPlainTextCommand
  | ReplaceParagraphPlainTextCommand
  | SetTextRunPropertiesCommand
  | ClearTextRunPropertiesCommand
  | SetParagraphPropertiesCommand
  | ClearParagraphPropertiesCommand
  | MoveShapeCommand
  | ResizeShapeCommand
  | SetShapeTransformCommand
  | SetShapeFillCommand
  | SetShapeOutlineCommand
  | AddTextBoxCommand
  | AddConnectorCommand
  | DeleteShapeCommand
  | GroupShapesCommand
  | UngroupShapeCommand
  | ReplaceImageCommand
  | SetPictureCropCommand
  | ClearPictureCropCommand
  | UpdateChartDataCommand
  | AddEmptySlideFromLayoutCommand
  | DuplicateSlideCommand
  | MoveSlideCommand
  | DeleteSlideCommand;

export type EditorOperationErrorCode =
  | "invalid-command"
  | "invalid-selection"
  | "empty-undo-stack"
  | "empty-redo-stack";

export interface EditorOperationFailure<
  Code extends EditorOperationErrorCode = EditorOperationErrorCode,
> {
  readonly ok: false;
  readonly code: Code;
  readonly message: string;
  readonly cause?: unknown;
}

export type EditorApplyCommandResult =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
      readonly warnings?: readonly EditorCommandWarning[];
    }
  | EditorOperationFailure<"invalid-command">;

export interface EditorCommandWarning {
  readonly code: "shared-media-part";
  readonly message: string;
  readonly mediaPartPath: string;
  readonly referenceCount: number;
}

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

  apply(command: EditorCommand): EditorApplyCommandResult {
    return this.applyAll([command]);
  }

  applyAll(commands: readonly EditorCommand[]): EditorApplyCommandResult {
    const before = this.#document;
    const beforeSelection = this.#selection;
    if (commands.length === 0) return { ok: true, document: before };

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
      return invalidCommandFailure(cause);
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
      (command) => command.kind === "groupShapes" || command.kind === "ungroupShape",
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

export function createEditorSession(document: PptxSourceModel): EditorSession {
  return new EditorSession(document);
}

type ApplyCommandAttempt =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
    }
  | EditorOperationFailure<"invalid-command">;

const EDITOR_COMMAND_KINDS: ReadonlySet<string> = new Set([
  "replaceTextRunPlainText",
  "replaceParagraphPlainText",
  "setTextRunProperties",
  "clearTextRunProperties",
  "setParagraphProperties",
  "clearParagraphProperties",
  "moveShape",
  "resizeShape",
  "setShapeTransform",
  "setShapeFill",
  "setShapeOutline",
  "addTextBox",
  "addConnector",
  "deleteShape",
  "groupShapes",
  "ungroupShape",
  "replaceImage",
  "setPictureCrop",
  "clearPictureCrop",
  "updateChartData",
  "addEmptySlideFromLayout",
  "duplicateSlide",
  "moveSlide",
  "deleteSlide",
]);

const EXPECTED_COMMAND_REJECTION_PREFIXES = [
  "replaceTextRunPlainText:",
  "replaceParagraphPlainText:",
  "setTextRunProperties:",
  "clearTextRunProperties:",
  "setParagraphProperties:",
  "clearParagraphProperties:",
  "moveShape:",
  "resizeShape:",
  "setShapeTransform:",
  "setShapeFill:",
  "setShapeOutline:",
  "addTextBox:",
  "addConnector:",
  "deleteShape:",
  "groupShapes:",
  "ungroupShape:",
  "replaceImageBytes:",
  "setPictureCrop:",
  "clearPictureCrop:",
  "updateChartData:",
  "addEmptySlideFromLayout:",
  "duplicateSlide:",
  "moveSlide:",
  "deleteSlide:",
  "updateTextRunProperties:",
  "updateParagraphProperties:",
  "updateShapeTransform:",
] as const;

function applyCommandToDocument(
  document: PptxSourceModel,
  command: EditorCommand,
): ApplyCommandAttempt {
  if (!EDITOR_COMMAND_KINDS.has(command.kind)) {
    throw new TypeError(`EditorSession: unsupported command kind '${String(command.kind)}'`);
  }
  switch (command.kind) {
    case "replaceTextRunPlainText":
    case "replaceParagraphPlainText":
    case "setTextRunProperties":
    case "clearTextRunProperties":
    case "setParagraphProperties":
    case "clearParagraphProperties":
    case "moveShape":
    case "resizeShape":
    case "setShapeTransform":
    case "setShapeFill":
    case "setShapeOutline":
    case "addTextBox":
    case "addConnector":
    case "deleteShape":
    case "groupShapes":
    case "ungroupShape":
    case "replaceImage":
    case "setPictureCrop":
    case "clearPictureCrop":
    case "updateChartData":
    case "addEmptySlideFromLayout":
    case "duplicateSlide":
    case "moveSlide":
    case "deleteSlide":
      return attemptCommand(() => executeCommand(document, command));
  }
}

function attemptCommand(operation: () => PptxSourceModel): ApplyCommandAttempt {
  try {
    return { ok: true, document: operation() };
  } catch (cause) {
    return invalidCommandFailure(cause);
  }
}

function invalidCommandFailure(cause: unknown): EditorOperationFailure<"invalid-command"> {
  if (
    !(cause instanceof Error) ||
    !EXPECTED_COMMAND_REJECTION_PREFIXES.some((prefix) => cause.message.startsWith(prefix))
  ) {
    throw cause;
  }
  return {
    ok: false,
    code: "invalid-command",
    message: cause.message,
    cause,
  };
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

function executeCommand(document: PptxSourceModel, command: EditorCommand): PptxSourceModel {
  switch (command.kind) {
    case "replaceTextRunPlainText":
      return replaceTextRunPlainTextCommand(document, command);
    case "replaceParagraphPlainText":
      return replaceParagraphPlainTextCommand(document, command);
    case "setTextRunProperties":
      return setTextRunPropertiesCommand(document, command);
    case "clearTextRunProperties":
      return clearTextRunPropertiesCommand(document, command);
    case "setParagraphProperties":
      return setParagraphPropertiesCommand(document, command);
    case "clearParagraphProperties":
      return clearParagraphPropertiesCommand(document, command);
    case "moveShape":
      return moveShape(document, command);
    case "resizeShape":
      return resizeShape(document, command);
    case "setShapeTransform":
      return setShapeTransform(document, command);
    case "setShapeFill":
      return setShapeFillCommand(document, command);
    case "setShapeOutline":
      return setShapeOutlineCommand(document, command);
    case "addTextBox":
      return addTextBoxCommand(document, command);
    case "addConnector":
      return addConnectorCommand(document, command);
    case "deleteShape":
      return deleteShape(document, command.handle);
    case "groupShapes":
      return groupShapesCommand(document, command);
    case "ungroupShape":
      return ungroupShapeCommand(document, command);
    case "replaceImage":
      return replaceImageBytes(document, command.handle, command.bytes);
    case "setPictureCrop":
      return setPictureCrop(document, command.handle, command);
    case "clearPictureCrop":
      return clearPictureCrop(document, command.handle);
    case "updateChartData":
      return updateChartDataCommand(document, command);
    case "addEmptySlideFromLayout":
      return addEmptySlideFromLayout(document, command);
    case "duplicateSlide":
      return duplicateSlide(document, command.handle);
    case "moveSlide":
      return moveSlide(document, command.handle, command);
    case "deleteSlide":
      return deleteSlide(document, command.handle);
  }
}

function groupShapesCommand(
  document: PptxSourceModel,
  command: GroupShapesCommand,
): PptxSourceModel {
  for (const handle of command.shapeHandles) {
    const shape = findCommandShape(document, handle, "groupShapes");
    if (shape !== undefined && !isGroupableSourceShape(shape)) {
      throw new Error(`groupShapes: shape kind '${shape.kind}' is not supported`);
    }
  }
  return groupShapes(document, command.shapeHandles);
}

function ungroupShapeCommand(
  document: PptxSourceModel,
  command: UngroupShapeCommand,
): PptxSourceModel {
  const group = findCommandShape(document, command.groupHandle, "ungroupShape");
  if (group?.kind === "group" && group.children.length === 0) {
    throw new Error("ungroupShape: group must contain at least one child");
  }
  return ungroupShape(document, command.groupHandle);
}

function findCommandShape(
  document: PptxSourceModel,
  handle: SourceHandle,
  operation: "groupShapes" | "ungroupShape",
): SourceShapeNode | undefined {
  try {
    return findShapeNodeBySourceHandle(document, handle);
  } catch (cause) {
    if (cause instanceof Error && cause.message.startsWith("findShapeNodeBySourceHandle:")) {
      throw new Error(`${operation}: ${cause.message}`, { cause });
    }
    throw cause;
  }
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

function updateChartDataCommand(
  document: PptxSourceModel,
  command: UpdateChartDataCommand,
): PptxSourceModel {
  if (!Array.isArray(command.series) || command.series.length === 0) {
    throw new Error("updateChartData: series must be a non-empty array");
  }
  for (const [index, series] of command.series.entries()) {
    if (!isObject(series)) {
      throw new Error(`updateChartData: series[${index}] must be an object`);
    }
    if (typeof series.name !== "string") {
      throw new Error(`updateChartData: series[${index}].name must be a string`);
    }
    if (!Array.isArray(series.categories) || !Array.isArray(series.values)) {
      throw new Error(`updateChartData: series[${index}] categories and values must be arrays`);
    }
  }
  return updateChartData(document, command.handle, command);
}

function replaceTextRunPlainTextCommand(
  document: PptxSourceModel,
  command: ReplaceTextRunPlainTextCommand,
): PptxSourceModel {
  if (typeof command.text !== "string") {
    throw new Error("replaceTextRunPlainText: text must be a string");
  }
  return replaceTextRunPlainText(document, command.handle, command.text);
}

function replaceParagraphPlainTextCommand(
  document: PptxSourceModel,
  command: ReplaceParagraphPlainTextCommand,
): PptxSourceModel {
  if (typeof command.text !== "string") {
    throw new Error("replaceParagraphPlainText: text must be a string");
  }
  return replaceParagraphPlainText(document, command.handle, command.text);
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
    if (edit === undefined || edit.sharedReferenceCount <= 1) continue;
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

function addTextBoxCommand(document: PptxSourceModel, command: AddTextBoxCommand): PptxSourceModel {
  requireFiniteEmu(command.offsetX, "addTextBox", "offsetX");
  requireFiniteEmu(command.offsetY, "addTextBox", "offsetY");
  requirePositiveFiniteEmu(command.width, "addTextBox", "width");
  requirePositiveFiniteEmu(command.height, "addTextBox", "height");
  if (command.text !== undefined && typeof command.text !== "string") {
    throw new Error("addTextBox: text must be a string");
  }
  if (command.name !== undefined && command.name.trim() === "") {
    throw new Error("addTextBox: name must be a non-empty string when provided");
  }
  return addTextBox(document, command.slideHandle, command);
}

function addConnectorCommand(
  document: PptxSourceModel,
  command: AddConnectorCommand,
): PptxSourceModel {
  requireFiniteEmu(command.offsetX, "addConnector", "offsetX");
  requireFiniteEmu(command.offsetY, "addConnector", "offsetY");
  requirePositiveFiniteEmu(command.width, "addConnector", "width");
  requirePositiveFiniteEmu(command.height, "addConnector", "height");
  return addConnector(document, command.slideHandle, command);
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

function setParagraphPropertiesCommand(
  document: PptxSourceModel,
  command: SetParagraphPropertiesCommand,
): PptxSourceModel {
  requireNonEmptyParagraphPropertySet(command.properties, "setParagraphProperties");
  validateParagraphPropertySet(command.properties, "setParagraphProperties");
  return setParagraphProperties(document, command.handle, command.properties);
}

function clearParagraphPropertiesCommand(
  document: PptxSourceModel,
  command: ClearParagraphPropertiesCommand,
): PptxSourceModel {
  if (command.properties.length === 0) {
    throw new Error("clearParagraphProperties: properties must contain at least one property name");
  }
  for (const property of command.properties) {
    if (!EDITABLE_PARAGRAPH_PROPERTY_SET.has(property)) {
      throw new Error(`clearParagraphProperties: unsupported paragraph property '${property}'`);
    }
  }
  return clearParagraphProperties(document, command.handle, command.properties);
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

function setShapeFillCommand(
  document: PptxSourceModel,
  command: SetShapeFillCommand,
): PptxSourceModel {
  validateShapeFill(command.fill, "setShapeFill");
  return setShapeFill(document, command.handle, command.fill);
}

function setShapeOutlineCommand(
  document: PptxSourceModel,
  command: SetShapeOutlineCommand,
): PptxSourceModel {
  validateShapeOutline(command.outline, "setShapeOutline");
  return setShapeOutline(document, command.handle, command.outline);
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

function requireNonEmptyParagraphPropertySet(
  properties: EditableParagraphProperties,
  commandName: "setParagraphProperties",
): void {
  if (Object.values(properties).every((value) => value === undefined)) {
    throw new Error(`${commandName}: properties must contain at least one defined property`);
  }
}

function validateParagraphPropertySet(
  properties: EditableParagraphProperties,
  commandName: "setParagraphProperties",
): void {
  for (const property of Object.keys(properties)) {
    if (!EDITABLE_PARAGRAPH_PROPERTY_SET.has(property)) {
      throw new Error(`${commandName}: unsupported paragraph property '${property}'`);
    }
  }
  if (properties.align !== undefined && !PARAGRAPH_ALIGN_VALUES.has(properties.align)) {
    throw new Error(`${commandName}: align must be left, center, right, or justify`);
  }
  if (
    properties.level !== undefined &&
    (!Number.isInteger(properties.level) || properties.level < 0 || properties.level > 8)
  ) {
    throw new Error(`${commandName}: level must be an integer from 0 to 8`);
  }
  if (properties.bullet !== undefined) {
    validateParagraphBullet(properties.bullet, commandName);
  }
}

function validateParagraphBullet(
  bullet: NonNullable<EditableParagraphProperties["bullet"]>,
  commandName: "setParagraphProperties",
): void {
  if (bullet.type === "none") return;
  if (bullet.type === "char") {
    if (bullet.char.length === 0) {
      throw new Error(`${commandName}: bullet.char must be a non-empty string`);
    }
    return;
  }
  if (bullet.type === "autoNum") {
    if (!AUTO_NUM_SCHEMES.has(bullet.scheme)) {
      throw new Error(`${commandName}: unsupported bullet auto-numbering scheme`);
    }
    if (!Number.isInteger(bullet.startAt) || bullet.startAt < 1) {
      throw new Error(`${commandName}: bullet.startAt must be a positive integer`);
    }
    return;
  }
  throw new Error(`${commandName}: unsupported bullet type`);
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

function validateShapeOutline(outline: EditableShapeOutline, commandName: "setShapeOutline"): void {
  if (outline.width === undefined && outline.fill === undefined) {
    throw new Error(`${commandName}: outline must set width or fill`);
  }
  if (outline.width !== undefined) {
    requirePositiveFiniteEmu(outline.width, commandName, "width");
  }
  if (outline.fill !== undefined) validateShapeFill(outline.fill, commandName);
}

function validateShapeFill(
  fill: EditableShapeFill,
  commandName: "setShapeFill" | "setShapeOutline",
): void {
  if (fill.kind === "none") return;
  if (fill.kind !== "solid") {
    throw new Error(`${commandName}: only solid and none fills are supported`);
  }
  if (fill.color.kind !== "srgb") {
    throw new Error(`${commandName}: only srgb solid fill colors are supported`);
  }
  if (!/^[0-9A-Fa-f]{6}$/.test(fill.color.hex)) {
    throw new Error(`${commandName}: color.hex must be a 6-digit hex value`);
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
  commandName:
    | "moveShape"
    | "resizeShape"
    | "setShapeTransform"
    | "setShapeOutline"
    | "addTextBox"
    | "addConnector",
  fieldName: string,
): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${commandName}: ${fieldName} must be a finite EMU value`);
  }
}

function requirePositiveFiniteEmu(
  value: Emu,
  commandName:
    | "moveShape"
    | "resizeShape"
    | "setShapeTransform"
    | "setShapeOutline"
    | "addTextBox"
    | "addConnector",
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
    if (edit.kind === "updateParagraphProperties") {
      const normalized = normalizeParagraphPropertiesEdit(edit, seenParagraphProperties);
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
    if (edit.kind === "updateShapeFill") {
      const key = editHandleNodeKey(edit);
      if (seenShapeFills.has(key)) {
        changed = true;
        continue;
      }
      seenShapeFills.add(key);
    }
    if (edit.kind === "updatePictureCrop") {
      const key = editHandleNodeKey(edit);
      if (seenPictureCrops.has(key)) {
        changed = true;
        continue;
      }
      seenPictureCrops.add(key);
    }
    if (edit.kind === "updateShapeOutline") {
      const normalized = normalizeShapeOutlineEdit(
        edit,
        seenShapeOutlineProperties,
        normalizedShapeOutlineEdits,
      );
      if (normalized === undefined) {
        changed = true;
        continue;
      }
      if (normalized.merged) {
        changed = true;
        continue;
      }
      if (!editorEditsEqual(normalized.edit, edit)) changed = true;
      normalizedReversed.push(normalized.edit);
      continue;
    }
    if (edit.kind === "updateChartData") {
      const key = sourceHandleKey(edit.handle);
      if (seenChartData.has(key)) {
        changed = true;
        continue;
      }
      seenChartData.add(key);
    }
    normalizedReversed.push(edit);
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
