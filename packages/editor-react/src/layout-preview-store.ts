import type {
  ConversionDiagnostic,
  PptxEditorTemplatePreviewResult,
  SourceHandle,
} from "pptx-glimpse";

export type LayoutPreviewLoader = (
  handle: SourceHandle,
) => Promise<PptxEditorTemplatePreviewResult>;

export type LayoutPreviewState =
  | { readonly status: "loading" }
  | { readonly status: "ready"; readonly svg: string }
  | { readonly status: "fallback"; readonly message: string };

interface QueuedPreview {
  readonly generation: number;
  readonly handle: SourceHandle;
  readonly key: string;
}

/** Session-local thumbnail cache and bounded async scheduler. */
export class LayoutPreviewStore {
  private readonly states = new Map<string, LayoutPreviewState>();
  private readonly queue: QueuedPreview[] = [];
  private readonly listeners = new Set<() => void>();
  private active = 0;
  private generation = 0;
  private version = 0;

  constructor(
    private readonly loadPreview: LayoutPreviewLoader,
    private readonly concurrency = 3,
  ) {
    if (!Number.isInteger(concurrency) || concurrency < 1) {
      throw new Error("LayoutPreviewStore concurrency must be a positive integer");
    }
  }

  readonly subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  readonly getVersion = (): number => this.version;

  get(handle: SourceHandle): LayoutPreviewState | undefined {
    return this.states.get(handleKey(handle));
  }

  load(handles: readonly SourceHandle[]): void {
    const generation = this.generation;
    let changed = false;
    for (const handle of handles) {
      const key = handleKey(handle);
      if (this.states.has(key)) continue;
      this.states.set(key, { status: "loading" });
      this.queue.push({ generation, handle, key });
      changed = true;
    }
    if (changed) this.notify();
    this.schedule();
  }

  dispose(): void {
    this.generation += 1;
    this.queue.length = 0;
    this.states.clear();
    this.listeners.clear();
  }

  private schedule(): void {
    while (this.active < this.concurrency) {
      const preview = this.queue.shift();
      if (preview === undefined) return;
      this.active += 1;
      void this.resolve(preview).finally(() => {
        this.active -= 1;
        this.schedule();
      });
    }
  }

  private async resolve(preview: QueuedPreview): Promise<void> {
    let state: LayoutPreviewState;
    try {
      const result = await this.loadPreview(preview.handle);
      state = previewState(result);
    } catch {
      state = { status: "fallback", message: "Preview failed" };
    }
    if (preview.generation !== this.generation) return;
    this.states.set(preview.key, state);
    this.notify();
  }

  private notify(): void {
    this.version += 1;
    for (const listener of this.listeners) listener();
  }
}

function previewState(result: PptxEditorTemplatePreviewResult): LayoutPreviewState {
  if (!result.ok) return { status: "fallback", message: "Preview unavailable" };
  if (result.diagnostics.some(isUnsupportedPreviewDiagnostic)) {
    return { status: "fallback", message: "Preview unsupported" };
  }
  return { status: "ready", svg: result.svg };
}

function isUnsupportedPreviewDiagnostic(diagnostic: ConversionDiagnostic): boolean {
  return diagnostic.severity === "error" || diagnostic.code.includes("unsupported");
}

function handleKey(handle: SourceHandle): string {
  return [
    handle.partPath ?? "",
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot === undefined ? "" : String(handle.orderingSlot),
  ].join("\u0000");
}
