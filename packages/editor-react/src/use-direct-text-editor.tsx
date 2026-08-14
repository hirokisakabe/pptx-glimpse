import {
  type EditorCommand,
  isPptxEditorError,
  type PptxEditorSession,
  type PptxEditorShapeBoundsPx,
  type PptxEditorShapeInfo,
  type SourceHandle,
} from "pptx-glimpse";
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";

import { DirectTextEditorLifecycle } from "./direct-text-editor-lifecycle.js";
import { viewBoxFromSvg } from "./editor-interaction-utils.js";
import type { PptxEditorController } from "./pptx-editor-controller.js";
import { shapeKey } from "./use-shape-transform-interactions.js";

interface DirectTextEditorRun {
  readonly key: string;
  readonly text: string;
  readonly handle?: SourceHandle;
}

interface DirectTextEditorParagraph {
  readonly key: string;
  readonly runs: readonly DirectTextEditorRun[];
}

interface DirectTextEditorState {
  readonly shapeKey: string;
  readonly bounds: PptxEditorShapeBoundsPx;
  readonly paragraphs: readonly DirectTextEditorParagraph[];
}

interface UseDirectTextEditorOptions {
  readonly controller: PptxEditorController<PptxEditorSession>;
  readonly currentIndex: number;
  readonly session: PptxEditorSession;
  readonly slideFrameRef: React.RefObject<HTMLDivElement | null>;
}

export function useDirectTextEditor({
  controller,
  currentIndex,
  session,
  slideFrameRef,
}: UseDirectTextEditorOptions) {
  const [editor, setEditor] = useState<DirectTextEditorState | null>(null);
  const editorRef = useRef<HTMLDivElement | null>(null);
  const editorStateRef = useRef<DirectTextEditorState | null>(null);
  const [lifecycle] = useState(() => new DirectTextEditorLifecycle());
  const compositionRef = useRef(false);
  const commitAfterCompositionRef = useRef(false);

  const close = useCallback(
    (restoreFocus = true) => {
      lifecycle.cancelComposition();
      editorStateRef.current = null;
      compositionRef.current = false;
      commitAfterCompositionRef.current = false;
      setEditor(null);
      if (restoreFocus) {
        window.setTimeout(() => slideFrameRef.current?.focus({ preventScroll: true }), 0);
      }
    },
    [lifecycle, slideFrameRef],
  );

  useLayoutEffect(() => {
    lifecycle.cancelComposition();
    editorStateRef.current = null;
    compositionRef.current = false;
    commitAfterCompositionRef.current = false;
    setEditor(null);
    return () => {
      lifecycle.invalidate();
      editorStateRef.current = null;
      compositionRef.current = false;
      commitAfterCompositionRef.current = false;
    };
  }, [controller, lifecycle]);

  const commit = useCallback(
    (restoreFocus = true): Promise<boolean> | undefined => {
      const activeEditor = editorStateRef.current;
      const editorElement = editorRef.current;
      if (activeEditor === null || editorElement === null) return;
      const currentCommit = lifecycle.currentCommit();
      if (currentCommit !== null) return currentCommit;
      const generation = lifecycle.currentGeneration();

      const commands = activeEditor.paragraphs.flatMap((paragraph) =>
        paragraph.runs.flatMap((run) => {
          if (run.handle === undefined) return [];
          const runElement = editorElement.querySelector<HTMLElement>(
            `[data-text-run-key="${cssAttributeValue(run.key)}"]`,
          );
          const text = runElement?.textContent ?? run.text;
          return text === run.text
            ? []
            : [
                {
                  kind: "replaceTextRunPlainText",
                  handle: run.handle,
                  text,
                } satisfies EditorCommand,
              ];
        }),
      );

      if (commands.length === 0) {
        close(restoreFocus);
        controller.setMessage("Text unchanged");
        return Promise.resolve(true);
      }

      const pendingCommit = (async () => {
        const committed = await controller.run(
          async () => {
            const result = await session.applyAll(commands);
            return commandMessage("Text updated", result.warnings);
          },
          {
            success: "Text updated",
            preferredIndex: currentIndex,
            recoverError: (error) =>
              isPptxEditorError(error) && error.code === "render-failed"
                ? "Text updated; slide preview could not refresh"
                : undefined,
          },
        );
        try {
          if (!lifecycle.isCurrent(generation)) return false;
          if (committed) {
            close(restoreFocus);
            return true;
          }
          return false;
        } finally {
          lifecycle.clearCommit(generation);
        }
      })();
      if (!lifecycle.setCommit(generation, pendingCommit)) {
        return lifecycle.currentCommit() ?? Promise.resolve(false);
      }
      return pendingCommit;
    },
    [close, controller, currentIndex, lifecycle, session],
  );

  const start = useCallback(
    (shape: PptxEditorShapeInfo) => {
      if (
        controller.getSnapshot().busy ||
        editorStateRef.current !== null ||
        shape.handle === undefined ||
        shape.bounds === undefined ||
        shape.textBody === undefined
      ) {
        return;
      }
      const paragraphs = shape.textBody.paragraphs.map((paragraph, paragraphIndex) => ({
        key: `paragraph-${paragraphIndex.toString()}`,
        runs: paragraph.runs.map((run, runIndex) => ({
          key: `run-${paragraphIndex.toString()}-${runIndex.toString()}`,
          text: run.text,
          ...(run.handle !== undefined ? { handle: run.handle } : {}),
        })),
      }));
      if (!paragraphs.some((paragraph) => paragraph.runs.some((run) => run.handle !== undefined))) {
        return;
      }

      controller.selectShape(shape.handle);
      const nextEditor = { shapeKey: shapeKey(shape), bounds: shape.bounds, paragraphs };
      editorStateRef.current = nextEditor;
      setEditor(nextEditor);
      controller.setMessage("Editing text");
    },
    [controller],
  );

  useEffect(() => {
    if (editor === null) return;
    const firstRun = editorRef.current?.querySelector<HTMLElement>(
      '[data-text-run-editable="true"]',
    );
    if (firstRun === undefined || firstRun === null) return;
    firstRun.focus();
    selectElementContents(firstRun);
  }, [editor]);

  const commitPendingEdits = useCallback(async () => {
    if (compositionRef.current) {
      const compositionCompletion = lifecycle.compositionPromise();
      if (compositionCompletion === undefined || !(await compositionCompletion)) return false;
    }
    const pendingCommit =
      lifecycle.currentCommit() ?? (editorStateRef.current === null ? undefined : commit(false));
    return pendingCommit === undefined || (await pendingCommit);
  }, [commit, lifecycle]);

  const waitForCurrentCommit = useCallback(async () => {
    const pendingCommit = lifecycle.currentCommit();
    return pendingCommit === null || (await pendingCommit);
  }, [lifecycle]);

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229) return;
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        close();
        controller.setMessage("Text edit canceled");
        return;
      }
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        void commit();
      }
    },
    [close, commit, controller],
  );

  const handleBlur = useCallback(
    (event: React.FocusEvent<HTMLDivElement>) => {
      const nextTarget = event.relatedTarget;
      if (nextTarget instanceof Node && event.currentTarget.contains(nextTarget)) return;
      if (compositionRef.current) {
        commitAfterCompositionRef.current = true;
        return;
      }
      void commit(false);
    },
    [commit],
  );

  const handleCompositionStart = useCallback(() => {
    compositionRef.current = true;
    commitAfterCompositionRef.current = false;
    lifecycle.beginComposition();
  }, [lifecycle]);

  const handleCompositionEnd = useCallback(() => {
    compositionRef.current = false;
    lifecycle.completeComposition();
    if (commitAfterCompositionRef.current && !editorRef.current?.contains(document.activeElement)) {
      commitAfterCompositionRef.current = false;
      void commit(false);
    }
  }, [commit, lifecycle]);

  return {
    commitPendingEdits,
    editor,
    editorRef,
    handleBlur,
    handleCompositionEnd,
    handleCompositionStart,
    handleKeyDown,
    isEditing: editor !== null,
    save: commit,
    start,
    waitForCurrentCommit,
  };
}

interface DirectTextEditorOverlayProps {
  readonly editor: DirectTextEditorState;
  readonly editorRef: React.RefObject<HTMLDivElement | null>;
  readonly slideSvg: string;
  readonly onBlur: (event: React.FocusEvent<HTMLDivElement>) => void;
  readonly onCompositionEnd: () => void;
  readonly onCompositionStart: () => void;
  readonly onDone: () => void;
  readonly onKeyDown: (event: React.KeyboardEvent<HTMLDivElement>) => void;
}

export function DirectTextEditorOverlay({
  editor,
  editorRef,
  slideSvg,
  onBlur,
  onCompositionEnd,
  onCompositionStart,
  onDone,
  onKeyDown,
}: DirectTextEditorOverlayProps) {
  return (
    <div
      ref={editorRef}
      className="direct-text-editor"
      data-shape-key={editor.shapeKey}
      data-testid="direct-text-editor"
      style={directTextEditorStyle(editor.bounds, viewBoxFromSvg(slideSvg))}
      onBlur={onBlur}
      onCompositionEnd={onCompositionEnd}
      onCompositionStart={onCompositionStart}
      onKeyDown={onKeyDown}
    >
      <div className="direct-text-editor-content">
        {editor.paragraphs.map((paragraph) => (
          <div
            className="direct-text-editor-paragraph"
            data-testid="direct-text-editor-paragraph"
            key={paragraph.key}
          >
            {paragraph.runs.map((run) => (
              <span
                className="direct-text-editor-run"
                contentEditable={run.handle === undefined ? undefined : "plaintext-only"}
                data-text-run-editable={run.handle === undefined ? undefined : "true"}
                data-text-run-key={run.key}
                data-testid="direct-text-editor-run"
                key={run.key}
                role={run.handle === undefined ? undefined : "textbox"}
                spellCheck={false}
                suppressContentEditableWarning
                onBeforeInput={(event) => {
                  const inputType = event.nativeEvent.inputType;
                  if (inputType === "insertParagraph" || inputType === "insertLineBreak") {
                    event.preventDefault();
                  }
                }}
                onPaste={(event) => {
                  event.preventDefault();
                  insertTextAtSelection(
                    event.clipboardData.getData("text/plain").replace(/\r?\n/g, " "),
                  );
                }}
              >
                {run.text}
              </span>
            ))}
          </div>
        ))}
      </div>
      <button
        className="direct-text-editor-done"
        data-testid="direct-text-editor-done"
        type="button"
        onClick={onDone}
      >
        Done
      </button>
    </div>
  );
}

function directTextEditorStyle(
  bounds: PptxEditorShapeBoundsPx,
  viewBox: string,
): React.CSSProperties {
  const [minX = 0, minY = 0, width = 960, height = 540] = viewBox
    .trim()
    .split(/[\s,]+/)
    .map(Number);
  const safeWidth = Number.isFinite(width) && width > 0 ? width : 960;
  const safeHeight = Number.isFinite(height) && height > 0 ? height : 540;
  return {
    left: `${(((bounds.x - minX) / safeWidth) * 100).toString()}%`,
    top: `${(((bounds.y - minY) / safeHeight) * 100).toString()}%`,
    width: `${((bounds.width / safeWidth) * 100).toString()}%`,
    height: `${((bounds.height / safeHeight) * 100).toString()}%`,
  };
}

function cssAttributeValue(value: string): string {
  return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function selectElementContents(element: HTMLElement): void {
  const selection = window.getSelection();
  if (selection === null) return;
  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
}

function insertTextAtSelection(text: string): void {
  const selection = window.getSelection();
  if (selection === null || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  range.deleteContents();
  const node = document.createTextNode(text);
  range.insertNode(node);
  range.setStartAfter(node);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
}

function commandMessage(
  fallback: string,
  warnings?: readonly {
    readonly code: string;
    readonly referenceCount?: number;
    readonly mediaPartPath?: string;
  }[],
): string {
  const sharedMedia = warnings?.find((warning) => warning.code === "shared-media-part");
  if (sharedMedia !== undefined) {
    return `${fallback}; shared media affects ${String(sharedMedia.referenceCount)} pictures: ${
      sharedMedia.mediaPartPath ?? "media part"
    }`;
  }
  return fallback;
}
