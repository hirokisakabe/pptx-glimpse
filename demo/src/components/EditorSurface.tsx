"use client";

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import {
  isPptxEditorError,
  type PptxEditorShapeBoundsPx,
  type PptxEditorShapeInfo,
  type PptxEditorSlideSvg,
  type EditorCommand,
  type PptxEditorSession,
  type SourceHandle,
} from "pptx-glimpse";

import { EditorSlideStrip } from "./EditorSlideStrip";
import { EditorHistoryToolbar, EditorToolbar, type EditorTextRunOption } from "./EditorToolbar";
import { DirectTextEditorLifecycle } from "./direct-text-editor-lifecycle";
import { type PreferredSlideIndex, useEditorController } from "./use-editor-controller";

const EMU_PER_PIXEL = 9525;
const MIN_SHAPE_SIZE = 8;
const MAX_IMAGE_REPLACEMENT_BYTES = 5 * 1024 * 1024;

type EditorSession = PptxEditorSession;
type ShapeTransformCommand = Extract<EditorCommand, { readonly kind: "setShapeTransform" }>;
type TextRunProperties = Extract<
  EditorCommand,
  { readonly kind: "setTextRunProperties" }
>["properties"];
type ClearTextRunProperties = Extract<
  EditorCommand,
  { readonly kind: "clearTextRunProperties" }
>["properties"];

interface EditorSurfaceProps {
  readonly editor: EditorSession;
  readonly children?: (controls: EditorSurfaceHostControls) => ReactNode;
}

export interface EditorSurfaceHostControls {
  readonly busy: boolean;
  readonly dirty: boolean;
  readonly message: string;
  readonly commitPendingEdits: () => Promise<boolean>;
  readonly hasUnsavedChanges: () => boolean;
  readonly markSaved: (
    history: ReturnType<PptxEditorSession["save"]>["history"],
    message: string,
  ) => boolean;
  readonly save: () => Promise<ReturnType<PptxEditorSession["save"]> | undefined>;
  readonly setError: (error: string) => void;
}

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

interface DragState {
  readonly interactionScope: object;
  readonly kind: "move" | "resize";
  readonly handle?: ResizeHandle;
  readonly shapeHandle: SourceHandle;
  readonly pointerId: number;
  readonly startPoint: Point;
  readonly startBounds: PptxEditorShapeBoundsPx;
}

interface Point {
  readonly x: number;
  readonly y: number;
}

type ResizeHandle = "nw" | "ne" | "sw" | "se";

/** Reusable editing UI. Document loading, download policy, and site navigation stay in its host. */
export function EditorSurface({ editor, children }: EditorSurfaceProps) {
  const {
    controller,
    slides,
    currentIndex,
    shapes: shapeOptions,
    selectedShapeKey,
    history,
    busy,
    dirty,
    message,
    error: operationError,
  } = useEditorController(editor);
  const [draftBounds, setDraftBounds] = useState<PptxEditorShapeBoundsPx | null>(null);
  const [directTextEditor, setDirectTextEditor] = useState<DirectTextEditorState | null>(null);
  const overlayRef = useRef<SVGSVGElement | null>(null);
  const slideFrameRef = useRef<HTMLDivElement | null>(null);
  const directTextEditorRef = useRef<HTMLDivElement | null>(null);
  const directTextEditorStateRef = useRef<DirectTextEditorState | null>(null);
  const [directTextLifecycle] = useState(() => new DirectTextEditorLifecycle());
  const dragStateRef = useRef<DragState | null>(null);
  const committedInteractionScopeRef = useRef<object | null>(null);
  const compositionRef = useRef(false);
  const commitAfterCompositionRef = useRef(false);

  const currentSlide = slides[currentIndex];
  const selectedShape = useMemo(() => {
    if (selectedShapeKey === null) return null;
    const shape = shapeOptions.find((candidate) => shapeKey(candidate) === selectedShapeKey);
    if (shape === undefined) return null;
    return draftBounds === null ? shape : { ...shape, bounds: draftBounds };
  }, [draftBounds, selectedShapeKey, shapeOptions]);

  const textRuns = useMemo<EditorTextRunOption[]>(() => {
    const sourceShapes = selectedShape === null ? [] : [selectedShape];
    return sourceShapes.flatMap((shape) =>
      (shape.textRuns ?? []).map((run, index) => ({
        label: `${shape.name ?? shape.kind} / run ${(index + 1).toString()}`,
        text: run.text,
        handle: run.handle,
      })),
    );
  }, [selectedShape]);

  useEffect(() => {
    setDraftBounds(null);
  }, [currentIndex, selectedShapeKey]);

  const runEditorOperation = useCallback(
    async (
      operation: (session: EditorSession) => Promise<string | void> | string | void,
      success: string,
      preferredIndex: PreferredSlideIndex<EditorSession> = currentIndex,
      historyAction: "mutation" | "undo" | "redo" = "mutation",
    ) => {
      const directTextCommit = directTextLifecycle.currentCommit();
      if (directTextCommit !== null && !(await directTextCommit)) return false;
      return controller.run(operation, { success, preferredIndex, historyAction });
    },
    [controller, currentIndex, directTextLifecycle],
  );

  const applyCommand = useCallback(
    (command: EditorCommand, success: string) =>
      runEditorOperation(async (session) => {
        const result = await session.apply(command);
        return commandMessage(success, result.warnings);
      }, success),
    [runEditorOperation],
  );

  const handleSelectShape = useCallback(
    (shape: PptxEditorShapeInfo, event?: React.PointerEvent<SVGRectElement>) => {
      if (controller.getSnapshot().busy || directTextEditorStateRef.current !== null) return;
      if (shape.handle === undefined) return;
      controller.selectShape(shape.handle);
      setDraftBounds(null);
      slideFrameRef.current?.focus({ preventScroll: true });
      if (event !== undefined && shape.editableTransform && shape.bounds !== undefined) {
        beginDrag(
          "move",
          undefined,
          shape.handle,
          controller,
          event,
          shape.bounds,
          dragStateRef,
          overlayRef,
        );
      }
    },
    [controller],
  );

  const updateDrag = useCallback(
    (event: PointerEvent) => {
      const dragState = dragStateRef.current;
      if (
        dragState === null ||
        committedInteractionScopeRef.current !== dragState.interactionScope ||
        event.pointerId !== dragState.pointerId
      ) {
        return;
      }
      const point = eventPoint(overlayRef.current, event.clientX, event.clientY);
      if (point === null) return;
      const dx = point.x - dragState.startPoint.x;
      const dy = point.y - dragState.startPoint.y;
      setDraftBounds(
        dragState.kind === "move"
          ? movedBounds(dragState.startBounds, dx, dy)
          : resizedBounds(dragState.startBounds, dragState.handle ?? "se", dx, dy),
      );
    },
    [controller],
  );

  const finishDrag = useCallback(
    async (event: PointerEvent) => {
      const dragState = dragStateRef.current;
      if (
        dragState === null ||
        committedInteractionScopeRef.current !== dragState.interactionScope ||
        event.pointerId !== dragState.pointerId
      ) {
        return;
      }
      dragStateRef.current = null;
      const point = eventPoint(overlayRef.current, event.clientX, event.clientY);
      if (point === null) {
        setDraftBounds(null);
        return;
      }

      const dx = point.x - dragState.startPoint.x;
      const dy = point.y - dragState.startPoint.y;
      const nextBounds =
        dragState.kind === "move"
          ? movedBounds(dragState.startBounds, dx, dy)
          : resizedBounds(dragState.startBounds, dragState.handle ?? "se", dx, dy);
      try {
        if (nextBounds === undefined || sameBounds(dragState.startBounds, nextBounds)) return;
        await applyCommand(
          {
            kind: "setShapeTransform",
            handle: dragState.shapeHandle,
            offsetX: pxToEmu(nextBounds.x),
            offsetY: pxToEmu(nextBounds.y),
            width: pxToEmu(nextBounds.width),
            height: pxToEmu(nextBounds.height),
          } satisfies ShapeTransformCommand,
          "Shape updated",
        );
      } finally {
        if (committedInteractionScopeRef.current === dragState.interactionScope) {
          setDraftBounds(null);
        }
      }
    },
    [applyCommand, controller],
  );

  useEffect(() => {
    const move = (event: PointerEvent) => updateDrag(event);
    const up = (event: PointerEvent) => {
      void finishDrag(event);
    };
    const cancelDrag = () => {
      dragStateRef.current = null;
      setDraftBounds(null);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", cancelDrag);
    window.addEventListener("blur", cancelDrag);
    return () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", cancelDrag);
      window.removeEventListener("blur", cancelDrag);
    };
  }, [finishDrag, updateDrag]);

  const handleResizeStart = useCallback(
    (handle: ResizeHandle, event: React.PointerEvent<SVGRectElement>) => {
      if (controller.getSnapshot().busy) return;
      if (selectedShape?.bounds === undefined || selectedShape.handle === undefined) return;
      event.preventDefault();
      event.stopPropagation();
      beginDrag(
        "resize",
        handle,
        selectedShape.handle,
        controller,
        event,
        selectedShape.bounds,
        dragStateRef,
        overlayRef,
      );
    },
    [controller, selectedShape],
  );

  const closeDirectTextEditor = useCallback(
    (restoreFocus = true) => {
      directTextLifecycle.cancelComposition();
      directTextEditorStateRef.current = null;
      compositionRef.current = false;
      commitAfterCompositionRef.current = false;
      setDirectTextEditor(null);
      if (restoreFocus) {
        window.setTimeout(() => slideFrameRef.current?.focus({ preventScroll: true }), 0);
      }
    },
    [directTextLifecycle],
  );

  useLayoutEffect(() => {
    committedInteractionScopeRef.current = controller;
    directTextLifecycle.cancelComposition();
    dragStateRef.current = null;
    directTextEditorStateRef.current = null;
    compositionRef.current = false;
    commitAfterCompositionRef.current = false;
    setDirectTextEditor(null);
    setDraftBounds(null);
    return () => {
      if (committedInteractionScopeRef.current === controller) {
        committedInteractionScopeRef.current = null;
      }
      directTextLifecycle.invalidate();
      dragStateRef.current = null;
      directTextEditorStateRef.current = null;
      compositionRef.current = false;
      commitAfterCompositionRef.current = false;
    };
  }, [controller, directTextLifecycle]);

  const commitDirectTextEditor = useCallback(
    (restoreFocus = true): Promise<boolean> | undefined => {
      const activeEditor = directTextEditorStateRef.current;
      const editorElement = directTextEditorRef.current;
      const session = editor;
      if (activeEditor === null || editorElement === null || session === null) return;
      const currentCommit = directTextLifecycle.currentCommit();
      if (currentCommit !== null) return currentCommit;
      const generation = directTextLifecycle.currentGeneration();

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
        closeDirectTextEditor(restoreFocus);
        controller.setMessage("Text unchanged");
        return Promise.resolve(true);
      }

      let commit!: Promise<boolean>;
      commit = (async () => {
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
          if (!directTextLifecycle.isCurrent(generation)) return false;
          if (committed) {
            closeDirectTextEditor(restoreFocus);
            return true;
          }
          return false;
        } finally {
          directTextLifecycle.clearCommit(generation, commit);
        }
      })();
      if (!directTextLifecycle.setCommit(generation, commit)) {
        return directTextLifecycle.currentCommit() ?? Promise.resolve(false);
      }
      return commit;
    },
    [closeDirectTextEditor, controller, currentIndex, directTextLifecycle, editor],
  );

  const startDirectTextEditor = useCallback(
    (shape: PptxEditorShapeInfo) => {
      if (
        controller.getSnapshot().busy ||
        directTextEditorStateRef.current !== null ||
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

      dragStateRef.current = null;
      controller.selectShape(shape.handle);
      setDraftBounds(null);
      const nextEditor = {
        shapeKey: shapeKey(shape),
        bounds: shape.bounds,
        paragraphs,
      };
      directTextEditorStateRef.current = nextEditor;
      setDirectTextEditor(nextEditor);
      controller.setMessage("Editing text");
    },
    [controller],
  );

  useEffect(() => {
    if (directTextEditor === null) return;
    const firstRun = directTextEditorRef.current?.querySelector<HTMLElement>(
      '[data-text-run-editable="true"]',
    );
    if (firstRun === undefined || firstRun === null) return;
    firstRun.focus();
    selectElementContents(firstRun);
  }, [directTextEditor]);

  const handleDirectTextEditorKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229) return;
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        closeDirectTextEditor();
        controller.setMessage("Text edit canceled");
        return;
      }
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        void commitDirectTextEditor();
      }
    },
    [closeDirectTextEditor, commitDirectTextEditor, controller],
  );

  const handleDirectTextEditorBlur = useCallback(
    (event: React.FocusEvent<HTMLDivElement>) => {
      const nextTarget = event.relatedTarget;
      if (nextTarget instanceof Node && event.currentTarget.contains(nextTarget)) return;
      if (compositionRef.current) {
        commitAfterCompositionRef.current = true;
        return;
      }
      void commitDirectTextEditor(false);
    },
    [commitDirectTextEditor],
  );

  const handleSlideFrameKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (
        event.key !== "Enter" ||
        event.target !== event.currentTarget ||
        selectedShape === null ||
        directTextEditorStateRef.current !== null
      ) {
        return;
      }
      event.preventDefault();
      startDirectTextEditor(selectedShape);
    },
    [selectedShape, startDirectTextEditor],
  );

  const handleApplyTextProperties = useCallback(
    (handle: SourceHandle, properties: TextRunProperties) =>
      applyCommand({ kind: "setTextRunProperties", handle, properties }, "Text style updated"),
    [applyCommand],
  );

  const handleClearTextProperties = useCallback(
    (handle: SourceHandle) => {
      const properties: ClearTextRunProperties = [
        "bold",
        "italic",
        "underline",
        "fontSize",
        "color",
        "typeface",
      ];
      return applyCommand(
        { kind: "clearTextRunProperties", handle, properties },
        "Text style cleared",
      );
    },
    [applyCommand],
  );

  const handleAddTextBox = useCallback(
    () =>
      runEditorOperation(async (session) => {
        await session.addTextBox(currentIndex + 1);
      }, "Text box added"),
    [currentIndex, runEditorOperation],
  );

  const handleDeleteShape = useCallback(() => {
    if (selectedShape?.handle === undefined) return;
    return applyCommand({ kind: "deleteShape", handle: selectedShape.handle }, "Shape deleted");
  }, [applyCommand, selectedShape]);

  const handleDuplicateSlide = useCallback(() => {
    if (currentSlide?.handle === undefined) return;
    const nextIndex = currentIndex + 1;
    return runEditorOperation(
      async (session) => {
        await session.apply({ kind: "duplicateSlide", handle: currentSlide.handle! });
      },
      "Slide duplicated",
      nextIndex,
    );
  }, [currentIndex, currentSlide, runEditorOperation]);

  const handleDeleteSlide = useCallback(() => {
    if (currentSlide?.handle === undefined || slides.length <= 1) return;
    const nextIndex = Math.max(0, currentIndex - 1);
    return runEditorOperation(
      async (session) => {
        await session.apply({ kind: "deleteSlide", handle: currentSlide.handle! });
      },
      "Slide deleted",
      nextIndex,
    );
  }, [currentIndex, currentSlide, runEditorOperation, slides.length]);

  const handleUndo = useCallback(() => {
    const selectedSlideHandle = currentSlide?.handle;
    return runEditorOperation(
      async (session) => {
        await session.undo();
      },
      "Undone",
      (session) => findSlideIndexByHandle(session.slides, selectedSlideHandle, currentIndex),
      "undo",
    );
  }, [currentIndex, currentSlide?.handle, runEditorOperation]);

  const handleRedo = useCallback(() => {
    const selectedSlideHandle = currentSlide?.handle;
    return runEditorOperation(
      async (session) => {
        await session.redo();
      },
      "Redone",
      (session) => findSlideIndexByHandle(session.slides, selectedSlideHandle, currentIndex),
      "redo",
    );
  }, [currentIndex, currentSlide?.handle, runEditorOperation]);

  const commitPendingEdits = useCallback(async () => {
    if (compositionRef.current) {
      const compositionCompletion = directTextLifecycle.compositionPromise();
      if (compositionCompletion === undefined || !(await compositionCompletion)) return false;
    }
    const commit =
      directTextLifecycle.currentCommit() ??
      (directTextEditorStateRef.current === null ? undefined : commitDirectTextEditor(false));
    return commit === undefined || (await commit);
  }, [commitDirectTextEditor, directTextLifecycle]);

  const handleSelectSlide = useCallback(
    async (index: number) => {
      if (await commitPendingEdits()) controller.selectSlide(index);
    },
    [commitPendingEdits, controller],
  );

  const handleMoveSlide = useCallback(
    async (fromIndex: number, toIndex: number) => {
      if (fromIndex === toIndex) return;
      const slide = slides[fromIndex];
      if (slide?.handle === undefined) return;
      const slideHandle = slide.handle;
      await runEditorOperation(
        async (session) => {
          await session.apply({ kind: "moveSlide", handle: slideHandle, toIndex });
        },
        `Slide ${slide.slideNumber.toString()} moved to position ${(toIndex + 1).toString()} of ${slides.length.toString()}`,
        (session) => findSlideIndexByHandle(session.slides, slideHandle, toIndex),
      );
    },
    [runEditorOperation, slides],
  );

  const handleImageReplacement = useCallback(
    async (file: File) => {
      if (!(await commitPendingEdits()) || selectedShape?.handle === undefined) return;
      const replacement = selectedShape.editableImageReplacement;
      if (replacement === undefined) return;
      if (file.size > MAX_IMAGE_REPLACEMENT_BYTES) {
        controller.setError("Replacement image must be 5 MB or smaller.");
        return;
      }
      if (file.type !== "" && file.type !== replacement.contentType) {
        controller.setError(`Replacement image must use ${replacement.contentType}.`);
        return;
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      const detectedContentType = detectImageContentType(bytes);
      if (detectedContentType !== replacement.contentType) {
        controller.setError(`Replacement image must use ${replacement.contentType}.`);
        return;
      }
      await applyCommand(
        {
          kind: "replaceImage",
          handle: selectedShape.handle,
          bytes,
        },
        "Image replaced",
      );
    },
    [applyCommand, commitPendingEdits, controller, selectedShape],
  );

  const hasUnsavedChanges = useCallback(
    () => controller.getSnapshot().dirty || directTextEditorStateRef.current !== null,
    [controller],
  );
  const markSaved = useCallback(
    (savedHistory: ReturnType<PptxEditorSession["save"]>["history"], savedMessage: string) =>
      controller.markSaved(savedHistory, savedMessage),
    [controller],
  );
  const save = useCallback(async () => {
    if (!(await commitPendingEdits()) || controller.getSnapshot().busy) return undefined;
    return controller.save();
  }, [commitPendingEdits, controller]);
  const setError = useCallback((error: string) => controller.setError(error), [controller]);

  const hostControls = useMemo<EditorSurfaceHostControls>(
    () => ({
      busy,
      dirty,
      message,
      commitPendingEdits,
      hasUnsavedChanges,
      markSaved,
      save,
      setError,
    }),
    [busy, commitPendingEdits, dirty, hasUnsavedChanges, markSaved, message, save, setError],
  );

  if (currentSlide === undefined) {
    return (
      <div className="loading" data-testid="editor-status">
        <div className="loading-mark" aria-hidden="true" />
        <p>{message}</p>
      </div>
    );
  }

  return (
    <section
      className="editor-workspace"
      aria-label="PPTX editor"
      data-testid="editor-workspace"
      data-editor-component="surface"
    >
      {children?.(hostControls)}
      <EditorHistoryToolbar
        busy={busy}
        history={history}
        onRedo={() => void handleRedo()}
        onUndo={() => void handleUndo()}
      />

      <div className="editor-shell">
        <EditorSlideStrip
          busy={busy}
          currentIndex={currentIndex}
          interactionScope={controller}
          slides={slides}
          onMove={(fromIndex, toIndex) => void handleMoveSlide(fromIndex, toIndex)}
          onSelect={(index) => void handleSelectSlide(index)}
        />

        <div className="editor-stage">
          <div
            ref={slideFrameRef}
            aria-label="Editable slide"
            className="editor-slide-frame"
            data-testid="editor-slide-frame"
            role="group"
            tabIndex={0}
            onKeyDown={handleSlideFrameKeyDown}
          >
            <div dangerouslySetInnerHTML={{ __html: currentSlide.svg }} />
            <svg
              ref={overlayRef}
              className={`editor-selection-overlay${directTextEditor === null ? "" : " editing"}`}
              data-testid="selection-overlay"
              viewBox={viewBoxFromSvg(currentSlide.svg)}
            >
              {shapeOptions.map((shape, index) => {
                if (shape.bounds === undefined) return null;
                return (
                  <rect
                    className="shape-hit-area"
                    data-editable-image-replacement={
                      shape.editableImageReplacement !== undefined ? "true" : undefined
                    }
                    data-editable-text={
                      shape.textBody?.paragraphs.some((paragraph) =>
                        paragraph.runs.some((run) => run.handle !== undefined),
                      )
                        ? "true"
                        : undefined
                    }
                    data-testid="shape-hit-area"
                    key={`${shapeKey(shape)}-${index.toString()}`}
                    x={shape.bounds.x}
                    y={shape.bounds.y}
                    width={shape.bounds.width}
                    height={shape.bounds.height}
                    onPointerDown={(event) => handleSelectShape(shape, event)}
                    onDoubleClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      startDirectTextEditor(shape);
                    }}
                  />
                );
              })}
              {selectedShape?.bounds !== undefined ? (
                <>
                  <rect
                    className="selection-box"
                    data-testid="selection-box"
                    x={selectedShape.bounds.x}
                    y={selectedShape.bounds.y}
                    width={selectedShape.bounds.width}
                    height={selectedShape.bounds.height}
                  />
                  {selectedShape.editableTransform && !busy && directTextEditor === null
                    ? (["nw", "ne", "sw", "se"] as const).map((handle) => {
                        const point = handlePoint(selectedShape.bounds!, handle);
                        return (
                          <rect
                            className={`selection-handle ${handle}`}
                            data-testid={`selection-handle-${handle}`}
                            key={handle}
                            x={point.x - 4}
                            y={point.y - 4}
                            width={8}
                            height={8}
                            onPointerDown={(event) => handleResizeStart(handle, event)}
                          />
                        );
                      })
                    : null}
                </>
              ) : null}
            </svg>
            {directTextEditor !== null ? (
              <div
                ref={directTextEditorRef}
                className="direct-text-editor"
                data-shape-key={directTextEditor.shapeKey}
                data-testid="direct-text-editor"
                style={directTextEditorStyle(
                  directTextEditor.bounds,
                  viewBoxFromSvg(currentSlide.svg),
                )}
                onBlur={handleDirectTextEditorBlur}
                onCompositionEnd={() => {
                  compositionRef.current = false;
                  directTextLifecycle.completeComposition();
                  if (
                    commitAfterCompositionRef.current &&
                    !directTextEditorRef.current?.contains(document.activeElement)
                  ) {
                    commitAfterCompositionRef.current = false;
                    void commitDirectTextEditor(false);
                  }
                }}
                onCompositionStart={() => {
                  compositionRef.current = true;
                  commitAfterCompositionRef.current = false;
                  directTextLifecycle.beginComposition();
                }}
                onKeyDown={handleDirectTextEditorKeyDown}
              >
                <div className="direct-text-editor-content">
                  {directTextEditor.paragraphs.map((paragraph) => (
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
                            if (
                              inputType === "insertParagraph" ||
                              inputType === "insertLineBreak"
                            ) {
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
                  onClick={() => void commitDirectTextEditor()}
                >
                  Done
                </button>
              </div>
            ) : null}
          </div>
        </div>

        <EditorToolbar
          busy={busy}
          error={operationError}
          selectionScope={controller}
          selectedShape={selectedShape}
          selectedShapeKey={selectedShapeKey}
          slidesCount={slides.length}
          textRuns={textRuns}
          onAddTextBox={() => void handleAddTextBox()}
          onApplyTextProperties={(handle, properties) =>
            void handleApplyTextProperties(handle, properties)
          }
          onClearTextProperties={(handle) => void handleClearTextProperties(handle)}
          onDeleteShape={() => void handleDeleteShape()}
          onDeleteSlide={() => void handleDeleteSlide()}
          onDuplicateSlide={() => void handleDuplicateSlide()}
          onReplaceImage={(file) => void handleImageReplacement(file)}
        />
      </div>
    </section>
  );
}

function beginDrag(
  kind: "move" | "resize",
  handle: ResizeHandle | undefined,
  shapeHandle: SourceHandle,
  interactionScope: object,
  event: React.PointerEvent<SVGRectElement>,
  startBounds: PptxEditorShapeBoundsPx,
  dragStateRef: React.MutableRefObject<DragState | null>,
  overlayRef: React.MutableRefObject<SVGSVGElement | null>,
) {
  const startPoint = eventPoint(overlayRef.current, event.clientX, event.clientY);
  if (startPoint === null) return;
  event.currentTarget.setPointerCapture(event.pointerId);
  dragStateRef.current = {
    interactionScope,
    kind,
    handle,
    shapeHandle,
    pointerId: event.pointerId,
    startPoint,
    startBounds,
  };
}

function eventPoint(svg: SVGSVGElement | null, clientX: number, clientY: number): Point | null {
  const matrix = svg?.getScreenCTM();
  if (svg === null || matrix === null || matrix === undefined) return null;
  const point = svg.createSVGPoint();
  point.x = clientX;
  point.y = clientY;
  const transformed = point.matrixTransform(matrix.inverse());
  return { x: transformed.x, y: transformed.y };
}

function movedBounds(
  bounds: PptxEditorShapeBoundsPx,
  dx: number,
  dy: number,
): PptxEditorShapeBoundsPx {
  return { ...bounds, x: bounds.x + dx, y: bounds.y + dy };
}

function resizedBounds(
  bounds: PptxEditorShapeBoundsPx,
  handle: ResizeHandle,
  dx: number,
  dy: number,
): PptxEditorShapeBoundsPx {
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  const next = { ...bounds };
  if (handle === "nw" || handle === "sw") {
    next.x = Math.min(bounds.x + dx, right - MIN_SHAPE_SIZE);
    next.width = right - next.x;
  }
  if (handle === "ne" || handle === "se") next.width = Math.max(MIN_SHAPE_SIZE, bounds.width + dx);
  if (handle === "nw" || handle === "ne") {
    next.y = Math.min(bounds.y + dy, bottom - MIN_SHAPE_SIZE);
    next.height = bottom - next.y;
  }
  if (handle === "sw" || handle === "se")
    next.height = Math.max(MIN_SHAPE_SIZE, bounds.height + dy);
  return next;
}

function handlePoint(bounds: PptxEditorShapeBoundsPx, handle: ResizeHandle): Point {
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  if (handle === "nw") return { x: bounds.x, y: bounds.y };
  if (handle === "ne") return { x: right, y: bounds.y };
  if (handle === "sw") return { x: bounds.x, y: bottom };
  return { x: right, y: bottom };
}

function sameBounds(a: PptxEditorShapeBoundsPx, b: PptxEditorShapeBoundsPx): boolean {
  return (
    Math.round(a.x) === Math.round(b.x) &&
    Math.round(a.y) === Math.round(b.y) &&
    Math.round(a.width) === Math.round(b.width) &&
    Math.round(a.height) === Math.round(b.height)
  );
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

function findSlideIndexByHandle(
  slides: readonly PptxEditorSlideSvg[],
  handle: SourceHandle | undefined,
  fallback: number,
): number {
  if (handle === undefined) return fallback;
  const index = slides.findIndex(
    (slide) => slide.handle !== undefined && handleKey(slide.handle) === handleKey(handle),
  );
  return index === -1 ? fallback : index;
}

function pxToEmu(value: number): ShapeTransformCommand["offsetX"] {
  return Math.round(value * EMU_PER_PIXEL) as ShapeTransformCommand["offsetX"];
}

function viewBoxFromSvg(svg: string): string {
  const fallback = svg.match(/\sviewBox="([^"]+)"/)?.[1] ?? "0 0 960 540";
  if (typeof DOMParser === "undefined") return fallback;
  const parsed = new DOMParser().parseFromString(svg, "image/svg+xml");
  const root = parsed.documentElement;
  const viewBox = root.getAttribute("viewBox");
  if (viewBox !== null && viewBox.trim() !== "") return viewBox;
  const width = parsePositiveSvgLength(root.getAttribute("width"));
  const height = parsePositiveSvgLength(root.getAttribute("height"));
  return width === undefined || height === undefined ? fallback : `0 0 ${width} ${height}`;
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

function parsePositiveSvgLength(value: string | null): number | undefined {
  if (value === null) return undefined;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function detectImageContentType(bytes: Uint8Array): string | undefined {
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return "image/png";
  }
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) return "image/jpeg";
  if (
    startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38, 0x37, 0x61]) ||
    startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38, 0x39, 0x61])
  ) {
    return "image/gif";
  }
  if (startsWithBytes(bytes, [0x42, 0x4d])) return "image/bmp";
  if (
    startsWithBytes(bytes, [0x49, 0x49, 0x2a, 0x00]) ||
    startsWithBytes(bytes, [0x4d, 0x4d, 0x00, 0x2a])
  ) {
    return "image/tiff";
  }
  if (
    startsWithBytes(bytes, [0x52, 0x49, 0x46, 0x46]) &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return "image/webp";
  }
  return undefined;
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((value, index) => bytes[index] === value);
}
