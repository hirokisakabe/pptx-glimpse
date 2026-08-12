"use client";

import { useMemo, useSyncExternalStore } from "react";
import type {
  PptxEditorSession,
  PptxEditorShapeInfo,
  PptxEditorSlideSvg,
  SourceHandle,
} from "pptx-glimpse";

type EditorHistory = PptxEditorSession["history"];
type EditorSaveResult = ReturnType<PptxEditorSession["save"]>;

export interface EditorControllerSession {
  readonly slides: readonly PptxEditorSlideSvg[];
  readonly history: EditorHistory;
  readonly selection: PptxEditorSession["selection"];
  shapes(slideNumber: number): readonly PptxEditorShapeInfo[];
  selectShape(handle: SourceHandle): void;
  save(): EditorSaveResult;
}

export interface EditorControllerState {
  readonly slides: readonly PptxEditorSlideSvg[];
  readonly currentIndex: number;
  readonly shapes: readonly PptxEditorShapeInfo[];
  readonly selectedShapeKey: string | null;
  readonly history: EditorHistory;
  readonly busy: boolean;
  readonly dirty: boolean;
  readonly message: string;
  readonly error: string;
}

export type PreferredSlideIndex<Session extends EditorControllerSession> =
  | number
  | ((session: Session) => number);

export interface EditorControllerOperationOptions<Session extends EditorControllerSession> {
  readonly success: string;
  readonly preferredIndex?: PreferredSlideIndex<Session>;
  readonly recoverError?: (error: unknown, session: Session) => string | undefined;
}

/**
 * Owns the state derived from one consumer-owned editor session. It deliberately has no DOM or
 * Next.js dependencies so the boundary can move to a browser React package in a later slice.
 */
export class EditorController<Session extends EditorControllerSession> {
  readonly session: Session;
  private state: EditorControllerState;
  private cleanUndoDepth: number;
  private operationQueue = Promise.resolve();
  private readonly listeners = new Set<() => void>();

  constructor(session: Session) {
    this.session = session;
    this.cleanUndoDepth = session.history.undoDepth;
    this.state = createEditorControllerSnapshot(session, 0, null, {
      busy: false,
      dirty: false,
      message: "",
      error: "",
    });
  }

  readonly subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  readonly getSnapshot = (): EditorControllerState => this.state;

  selectSlide(index: number): void {
    this.sync(index);
  }

  selectShape(handle: SourceHandle): void {
    this.session.selectShape(handle);
    this.update({
      ...this.state,
      selectedShapeKey: handleKey(handle),
    });
  }

  setMessage(message: string): void {
    this.update({ ...this.state, message });
  }

  setError(error: string): void {
    this.update({ ...this.state, error });
  }

  run(
    operation: (session: Session) => Promise<string | void> | string | void,
    options: EditorControllerOperationOptions<Session>,
  ): Promise<boolean> {
    return this.enqueue(async () => {
      this.update({ ...this.state, busy: true, error: "" });
      try {
        const message = await operation(this.session);
        const preferredIndex =
          typeof options.preferredIndex === "function"
            ? options.preferredIndex(this.session)
            : (options.preferredIndex ?? this.state.currentIndex);
        this.sync(preferredIndex, message ?? options.success);
        return true;
      } catch (error) {
        const recoveredMessage = options.recoverError?.(error, this.session);
        if (recoveredMessage !== undefined) {
          this.sync(this.state.currentIndex, recoveredMessage);
          return true;
        }
        this.update({
          ...this.state,
          history: this.session.history,
          dirty: this.isDirty(),
          error: errorMessage(error),
        });
        return false;
      } finally {
        this.update({ ...this.state, busy: false });
      }
    });
  }

  save(): Promise<EditorSaveResult | undefined> {
    return this.enqueue(async () => {
      this.update({ ...this.state, busy: true, error: "" });
      try {
        const saved = this.session.save();
        this.cleanUndoDepth = saved.history.undoDepth;
        this.update({
          ...this.state,
          history: saved.history,
          dirty: false,
          message: "PPTX downloaded",
        });
        return saved;
      } catch (error) {
        this.update({ ...this.state, error: errorMessage(error) });
        return undefined;
      } finally {
        this.update({ ...this.state, busy: false });
      }
    });
  }

  private sync(preferredIndex: number, message = this.state.message): void {
    this.update(
      createEditorControllerSnapshot(this.session, preferredIndex, this.state.selectedShapeKey, {
        busy: this.state.busy,
        dirty: this.isDirty(),
        message,
        error: this.state.error,
      }),
    );
  }

  private isDirty(): boolean {
    return this.session.history.undoDepth !== this.cleanUndoDepth;
  }

  private enqueue<Result>(operation: () => Promise<Result>): Promise<Result> {
    const result = this.operationQueue.then(operation, operation);
    this.operationQueue = result.then(
      () => undefined,
      () => undefined,
    );
    return result;
  }

  private update(state: EditorControllerState): void {
    if (state === this.state) return;
    this.state = state;
    for (const listener of this.listeners) listener();
  }
}

export function useEditorController(editor: PptxEditorSession): EditorControllerState & {
  readonly controller: EditorController<PptxEditorSession>;
} {
  const controller = useMemo(() => new EditorController(editor), [editor]);
  const state = useSyncExternalStore(
    controller.subscribe,
    controller.getSnapshot,
    controller.getSnapshot,
  );
  return useMemo(() => ({ ...state, controller }), [controller, state]);
}

export function createEditorControllerSnapshot(
  session: EditorControllerSession,
  preferredIndex: number,
  selectedShapeKey: string | null,
  status: Pick<EditorControllerState, "busy" | "dirty" | "message" | "error">,
): EditorControllerState {
  const slides = [...session.slides];
  const currentIndex = clamp(preferredIndex, 0, Math.max(slides.length - 1, 0));
  const shapes = session
    .shapes(currentIndex + 1)
    .filter((shape) => shape.handle !== undefined && shape.bounds !== undefined);
  const responseSelection = session.selection?.shapeHandle;
  const nextSelectionKey =
    responseSelection !== undefined
      ? handleKey(responseSelection)
      : selectedShapeKey !== null && shapes.some((shape) => shapeKey(shape) === selectedShapeKey)
        ? selectedShapeKey
        : null;

  return {
    slides,
    currentIndex,
    shapes: [...shapes],
    selectedShapeKey: nextSelectionKey,
    history: session.history,
    ...status,
  };
}

function shapeKey(shape: PptxEditorShapeInfo): string {
  return shape.handle === undefined ? "" : handleKey(shape.handle);
}

function handleKey(handle: SourceHandle): string {
  return [
    handle.partPath ?? "",
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot === undefined ? "" : String(handle.orderingSlot),
  ].join("\u0000");
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
