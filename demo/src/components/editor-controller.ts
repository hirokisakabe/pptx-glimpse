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
  readonly historyAction?: "mutation" | "undo" | "redo";
}

/**
 * Owns the state derived from one consumer-owned editor session. It deliberately has no DOM or
 * Next.js dependencies so the boundary can move to a browser React package in a later slice.
 */
export class EditorController<Session extends EditorControllerSession> {
  readonly session: Session;
  private state: EditorControllerState;
  private historyRevisions: number[];
  private cleanRevision: number;
  private nextRevision: number;
  private operationQueue = Promise.resolve();
  private pendingOperations = 0;
  private readonly listeners = new Set<() => void>();

  constructor(session: Session) {
    this.session = session;
    const historyLength = session.history.undoDepth + session.history.redoDepth + 1;
    this.historyRevisions = Array.from({ length: historyLength }, (_, index) => index);
    this.cleanRevision = this.historyRevisions[session.history.undoDepth] ?? 0;
    this.nextRevision = historyLength;
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

  selectSlide(index: number): boolean {
    if (this.pendingOperations > 0) return false;
    this.sync(index);
    return true;
  }

  selectShape(handle: SourceHandle): boolean {
    if (this.pendingOperations > 0) return false;
    this.session.selectShape(handle);
    this.update({
      ...this.state,
      selectedShapeKey: handleKey(handle),
    });
    return true;
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
      this.update({ ...this.state, error: "" });
      const beforeHistory = this.session.history;
      try {
        const message = await operation(this.session);
        this.reconcileHistory(options.historyAction ?? "mutation", beforeHistory);
        const preferredIndex =
          typeof options.preferredIndex === "function"
            ? options.preferredIndex(this.session)
            : (options.preferredIndex ?? this.state.currentIndex);
        this.sync(preferredIndex, message ?? options.success);
        return true;
      } catch (error) {
        this.reconcileHistory(options.historyAction ?? "mutation", beforeHistory);
        const recoveredMessage = options.recoverError?.(error, this.session);
        if (recoveredMessage !== undefined) {
          this.sync(this.state.currentIndex, recoveredMessage);
          this.update({ ...this.state, error: errorMessage(error) });
          return true;
        }
        this.update({
          ...this.state,
          history: this.session.history,
          dirty: this.isDirty(),
          error: errorMessage(error),
        });
        return false;
      }
    });
  }

  save(): Promise<EditorSaveResult | undefined> {
    return this.enqueue(async () => {
      this.update({ ...this.state, error: "" });
      try {
        const saved = this.session.save();
        this.update({
          ...this.state,
          history: saved.history,
        });
        return saved;
      } catch (error) {
        this.update({ ...this.state, error: errorMessage(error) });
        return undefined;
      }
    });
  }

  markSaved(history: EditorHistory, message: string): boolean {
    if (!sameHistory(history, this.session.history)) return false;
    this.cleanRevision = this.currentRevision();
    this.update({ ...this.state, history, dirty: false, message, error: "" });
    return true;
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
    return this.currentRevision() !== this.cleanRevision;
  }

  private enqueue<Result>(operation: () => Promise<Result>): Promise<Result> {
    this.pendingOperations += 1;
    if (this.pendingOperations === 1) this.update({ ...this.state, busy: true });
    const result = this.operationQueue.then(operation, operation);
    this.operationQueue = result.then(
      () => undefined,
      () => undefined,
    );
    return result.finally(() => {
      this.pendingOperations -= 1;
      if (this.pendingOperations === 0) this.update({ ...this.state, busy: false });
    });
  }

  private reconcileHistory(
    action: NonNullable<EditorControllerOperationOptions<Session>["historyAction"]>,
    before: EditorHistory,
  ): void {
    const after = this.session.history;
    if (sameHistory(before, after)) return;
    if (action === "mutation") {
      this.historyRevisions = this.historyRevisions.slice(0, before.undoDepth + 1);
      while (this.historyRevisions.length <= after.undoDepth) {
        this.historyRevisions.push(this.nextRevision);
        this.nextRevision += 1;
      }
      return;
    }
    while (this.historyRevisions.length <= after.undoDepth + after.redoDepth) {
      this.historyRevisions.push(this.nextRevision);
      this.nextRevision += 1;
    }
  }

  private currentRevision(): number {
    return this.historyRevisions[this.session.history.undoDepth] ?? -1;
  }

  private update(state: EditorControllerState): void {
    if (state === this.state) return;
    this.state = state;
    for (const listener of this.listeners) listener();
  }
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
  const responseSelectionKey =
    responseSelection === undefined ? null : handleKey(responseSelection);
  const nextSelectionKey =
    responseSelectionKey !== null &&
    shapes.some((shape) => shapeKey(shape) === responseSelectionKey)
      ? responseSelectionKey
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

function sameHistory(a: EditorHistory, b: EditorHistory): boolean {
  return (
    a.canUndo === b.canUndo &&
    a.canRedo === b.canRedo &&
    a.undoDepth === b.undoDepth &&
    a.redoDepth === b.redoDepth
  );
}
