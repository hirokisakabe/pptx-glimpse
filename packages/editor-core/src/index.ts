import {
  type Emu,
  findShapeNodeBySourceHandle,
  type PptxSourceModel,
  type PptxSourceModelEdit,
  type PptxSourceModelShapeTransformEdit,
  type PptxSourceModelTextRunEdit,
  replaceTextRunPlainText,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTransform,
  updateShapeTransform,
} from "@pptx-glimpse/document";

export interface ReplaceTextRunPlainTextCommand {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
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

export type EditorCommand = ReplaceTextRunPlainTextCommand | MoveShapeCommand | ResizeShapeCommand;

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

interface HistoryEntry {
  readonly before: PptxSourceModel;
  readonly after: PptxSourceModel;
}

export class EditorSession {
  #document: PptxSourceModel;
  readonly #undoStack: HistoryEntry[] = [];
  readonly #redoStack: HistoryEntry[] = [];

  constructor(document: PptxSourceModel) {
    this.#document = document;
  }

  get document(): PptxSourceModel {
    return this.#document;
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

  apply(command: EditorCommand): EditorApplyCommandResult {
    const before = this.#document;
    let after: PptxSourceModel;

    try {
      after = normalizeEditorEdits(applyCommandToDocument(before, command));
    } catch (error) {
      return {
        ok: false,
        code: "invalid-command",
        message: error instanceof Error ? error.message : "Editor command was rejected.",
        cause: error,
      };
    }

    this.#document = after;
    this.#undoStack.push({ before, after });
    this.#redoStack.length = 0;

    return { ok: true, document: after };
  }

  undo(): EditorHistoryResult {
    const entry = this.#undoStack.pop();
    if (entry === undefined) return { ok: false, reason: "empty-undo-stack" };

    this.#document = entry.before;
    this.#redoStack.push(entry);

    return { ok: true, document: entry.before };
  }

  redo(): EditorHistoryResult {
    const entry = this.#redoStack.pop();
    if (entry === undefined) return { ok: false, reason: "empty-redo-stack" };

    this.#document = entry.after;
    this.#undoStack.push(entry);

    return { ok: true, document: entry.after };
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
    case "moveShape":
      return moveShape(document, command);
    case "resizeShape":
      return resizeShape(document, command);
  }
}

function moveShape(document: PptxSourceModel, command: MoveShapeCommand): PptxSourceModel {
  const current = requireEditableShapeTransform(document, command.handle, "moveShape");
  return updateShapeTransform(document, command.handle, {
    offsetX: command.offsetX,
    offsetY: command.offsetY,
    width: current.width,
    height: current.height,
  });
}

function resizeShape(document: PptxSourceModel, command: ResizeShapeCommand): PptxSourceModel {
  const current = requireEditableShapeTransform(document, command.handle, "resizeShape");
  return updateShapeTransform(document, command.handle, {
    offsetX: current.offsetX,
    offsetY: current.offsetY,
    width: command.width,
    height: command.height,
  });
}

function requireEditableShapeTransform(
  document: PptxSourceModel,
  handle: SourceHandle,
  commandName: "moveShape" | "resizeShape",
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

function hasTransform(shape: SourceShapeNode): shape is SourceShapeNode & {
  readonly transform: SourceTransform;
} {
  return shape.kind !== "raw" && shape.transform !== undefined;
}

function normalizeEditorEdits(document: PptxSourceModel): PptxSourceModel {
  const edits = document.edits;
  if (edits === undefined) return document;

  const seenTextRuns = new Set<string>();
  const seenShapeTransforms = new Set<string>();
  const normalizedReversed: PptxSourceModelEdit[] = [];

  for (let index = edits.length - 1; index >= 0; index -= 1) {
    const edit = edits[index];
    if (edit.kind === "replaceTextRunPlainText") {
      const key = editHandleNodeKey(edit);
      if (seenTextRuns.has(key)) continue;
      seenTextRuns.add(key);
    }
    if (edit.kind === "updateShapeTransform") {
      const key = editHandleNodeKey(edit);
      if (seenShapeTransforms.has(key)) continue;
      seenShapeTransforms.add(key);
    }
    normalizedReversed.push(edit);
  }

  if (normalizedReversed.length === edits.length) return document;
  return {
    ...document,
    edits: normalizedReversed.reverse(),
  };
}

function editHandleNodeKey(
  edit: PptxSourceModelTextRunEdit | PptxSourceModelShapeTransformEdit,
): string {
  return [edit.handle.partPath, edit.handle.nodeId ?? "", edit.handle.relationshipId ?? ""].join(
    "\u0000",
  );
}
