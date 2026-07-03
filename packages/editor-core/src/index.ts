import {
  type PptxSourceModel,
  type PptxSourceModelEdit,
  type PptxSourceModelTextRunEdit,
  replaceTextRunPlainText,
  type SourceHandle,
} from "@pptx-glimpse/document";

export interface ReplaceTextRunPlainTextCommand {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

export type EditorCommand = ReplaceTextRunPlainTextCommand;

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
  }
}

function normalizeEditorEdits(document: PptxSourceModel): PptxSourceModel {
  const edits = document.edits;
  if (edits === undefined) return document;

  const seenTextRuns = new Set<string>();
  const normalizedReversed: PptxSourceModelEdit[] = [];

  for (let index = edits.length - 1; index >= 0; index -= 1) {
    const edit = edits[index];
    if (edit.kind === "replaceTextRunPlainText") {
      const key = textRunEditKey(edit);
      if (seenTextRuns.has(key)) continue;
      seenTextRuns.add(key);
    }
    normalizedReversed.push(edit);
  }

  if (normalizedReversed.length === edits.length) return document;
  return {
    ...document,
    edits: normalizedReversed.reverse(),
  };
}

function textRunEditKey(edit: PptxSourceModelTextRunEdit): string {
  return [edit.handle.partPath, edit.handle.nodeId ?? "", edit.handle.relationshipId ?? ""].join(
    "\u0000",
  );
}
