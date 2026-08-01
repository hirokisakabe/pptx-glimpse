"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  isPptxEditorError,
  type PptxEditorShapeBoundsPx,
  type PptxEditorShapeInfo,
  type PptxEditorSlideSvg,
  type EditorCommand,
  type PptxEditorSession,
  type SourceHandle,
} from "pptx-glimpse";

const EMU_PER_PIXEL = 9525;
const MIN_SHAPE_SIZE = 8;
const MAX_IMAGE_REPLACEMENT_BYTES = 5 * 1024 * 1024;
const SLIDE_DRAG_THRESHOLD_PX = 4;

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

interface EditorWorkspaceProps {
  readonly editor: EditorSession;
  readonly fileName: string;
  readonly fontFileCount: number;
  readonly onAddFonts: () => void;
  readonly onOpenPptx: () => void;
  readonly onOpenSample: () => void;
}

interface TextRunOption {
  readonly label: string;
  readonly text: string;
  readonly handle: SourceHandle;
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

interface SlideDropTarget {
  readonly index: number;
  readonly edge: "before" | "after";
}

interface SlideSortDragState {
  readonly fromIndex: number;
  readonly pointerId: number;
  readonly startPoint: Point;
  readonly hasMoved: boolean;
}

type PreferredSlideIndex = number | ((session: EditorSession) => number);

type ResizeHandle = "nw" | "ne" | "sw" | "se";

export function EditorWorkspace({
  editor,
  fileName,
  fontFileCount,
  onAddFonts,
  onOpenPptx,
  onOpenSample,
}: EditorWorkspaceProps) {
  const [downloadName, setDownloadName] = useState(() => fileStem(fileName));
  const [slides, setSlides] = useState<PptxEditorSlideSvg[]>(() => [...editor.slides]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [shapeOptions, setShapeOptions] = useState<PptxEditorShapeInfo[]>(() => [
    ...editor.shapes(1).filter((shape) => shape.handle && shape.bounds),
  ]);
  const [selectedShapeKey, setSelectedShapeKey] = useState<string | null>(null);
  const [draftBounds, setDraftBounds] = useState<PptxEditorShapeBoundsPx | null>(null);
  const [selectedRunIndex, setSelectedRunIndex] = useState(0);
  const [directTextEditor, setDirectTextEditor] = useState<DirectTextEditorState | null>(null);
  const [fontSize, setFontSize] = useState("24");
  const [typeface, setTypeface] = useState("");
  const [color, setColor] = useState("#2454a6");
  const [history, setHistory] = useState(editor.history);
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);
  const [operationError, setOperationError] = useState("");
  const [draggedSlideIndex, setDraggedSlideIndex] = useState<number | null>(null);
  const [slideDropTarget, setSlideDropTarget] = useState<SlideDropTarget | null>(null);
  const overlayRef = useRef<SVGSVGElement | null>(null);
  const slideFrameRef = useRef<HTMLDivElement | null>(null);
  const directTextEditorRef = useRef<HTMLDivElement | null>(null);
  const directTextEditorStateRef = useRef<DirectTextEditorState | null>(null);
  const directTextCommitPromiseRef = useRef<Promise<boolean> | null>(null);
  const dragStateRef = useRef<DragState | null>(null);
  const slideSortDragRef = useRef<SlideSortDragState | null>(null);
  const suppressSlideClickRef = useRef(false);
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const busyRef = useRef(false);
  const dirtyRef = useRef(false);
  const cleanUndoDepthRef = useRef(editor.history.undoDepth);
  const compositionRef = useRef(false);
  const commitAfterCompositionRef = useRef(false);

  const currentSlide = slides[currentIndex];
  const selectedShape = useMemo(() => {
    if (selectedShapeKey === null) return null;
    const shape = shapeOptions.find((candidate) => shapeKey(candidate) === selectedShapeKey);
    if (shape === undefined) return null;
    return draftBounds === null ? shape : { ...shape, bounds: draftBounds };
  }, [draftBounds, selectedShapeKey, shapeOptions]);

  const textRuns = useMemo<TextRunOption[]>(() => {
    const sourceShapes = selectedShape === null ? [] : [selectedShape];
    return sourceShapes.flatMap((shape) =>
      (shape.textRuns ?? []).map((run, index) => ({
        label: `${shape.name ?? shape.kind} / run ${(index + 1).toString()}`,
        text: run.text,
        handle: run.handle,
      })),
    );
  }, [selectedShape]);

  const selectedRun = textRuns[selectedRunIndex];

  useEffect(() => {
    setDownloadName(fileStem(fileName));
  }, [fileName]);

  const syncFromEditor = useCallback(
    (session: EditorSession, preferredIndex = currentIndex) => {
      const nextSlides = [...session.slides];
      const nextIndex = clamp(preferredIndex, 0, Math.max(nextSlides.length - 1, 0));
      const nextShapes = session
        .shapes(nextIndex + 1)
        .filter((shape) => shape.handle !== undefined && shape.bounds !== undefined);
      const responseSelection = session.selection?.shapeHandle;
      const nextSelectionKey =
        responseSelection !== undefined
          ? handleKey(responseSelection)
          : selectedShapeKey !== null &&
              nextShapes.some((shape) => shapeKey(shape) === selectedShapeKey)
            ? selectedShapeKey
            : null;

      setSlides(nextSlides);
      setCurrentIndex(nextIndex);
      setShapeOptions([...nextShapes]);
      setSelectedShapeKey(nextSelectionKey);
      setDraftBounds(null);
      setHistory(session.history);
      setSelectedRunIndex(0);
    },
    [currentIndex, selectedShapeKey],
  );

  useEffect(() => {
    setSelectedRunIndex(0);
  }, [selectedShapeKey]);

  useEffect(() => {
    syncFromEditor(editor, currentIndex);
  }, [currentIndex, editor, syncFromEditor]);

  const runEditorOperation = useCallback(
    async (
      operation: (session: EditorSession) => Promise<string | void> | string | void,
      success: string,
      preferredIndex: PreferredSlideIndex = currentIndex,
    ) => {
      const directTextCommit = directTextCommitPromiseRef.current;
      if (directTextCommit !== null && !(await directTextCommit)) return;
      if (busyRef.current) return;
      busyRef.current = true;
      setBusy(true);
      setOperationError("");
      try {
        const messageOverride = await operation(editor);
        const nextIndex =
          typeof preferredIndex === "function" ? preferredIndex(editor) : preferredIndex;
        syncFromEditor(editor, nextIndex);
        dirtyRef.current = editor.history.undoDepth !== cleanUndoDepthRef.current;
        setMessage(messageOverride ?? success);
      } catch (error) {
        setOperationError(error instanceof Error ? error.message : String(error));
      } finally {
        busyRef.current = false;
        setBusy(false);
      }
    },
    [currentIndex, editor, syncFromEditor],
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
      if (busyRef.current || directTextEditorStateRef.current !== null) return;
      if (shape.handle === undefined) return;
      editor?.selectShape(shape.handle);
      setSelectedShapeKey(shapeKey(shape));
      setDraftBounds(null);
      slideFrameRef.current?.focus({ preventScroll: true });
      if (event !== undefined && shape.editableTransform && shape.bounds !== undefined) {
        beginDrag("move", undefined, shape.handle, event, shape.bounds, dragStateRef, overlayRef);
      }
    },
    [editor],
  );

  const updateDrag = useCallback((event: PointerEvent) => {
    const dragState = dragStateRef.current;
    if (dragState === null || event.pointerId !== dragState.pointerId) return;
    const point = eventPoint(overlayRef.current, event.clientX, event.clientY);
    if (point === null) return;
    const dx = point.x - dragState.startPoint.x;
    const dy = point.y - dragState.startPoint.y;
    setDraftBounds(
      dragState.kind === "move"
        ? movedBounds(dragState.startBounds, dx, dy)
        : resizedBounds(dragState.startBounds, dragState.handle ?? "se", dx, dy),
    );
  }, []);

  const finishDrag = useCallback(
    async (event: PointerEvent) => {
      const dragState = dragStateRef.current;
      if (dragState === null || event.pointerId !== dragState.pointerId) return;
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
        setDraftBounds(null);
      }
    },
    [applyCommand],
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
      if (busyRef.current) return;
      if (selectedShape?.bounds === undefined || selectedShape.handle === undefined) return;
      event.preventDefault();
      event.stopPropagation();
      beginDrag(
        "resize",
        handle,
        selectedShape.handle,
        event,
        selectedShape.bounds,
        dragStateRef,
        overlayRef,
      );
    },
    [selectedShape],
  );

  const closeDirectTextEditor = useCallback((restoreFocus = true) => {
    directTextEditorStateRef.current = null;
    compositionRef.current = false;
    commitAfterCompositionRef.current = false;
    setDirectTextEditor(null);
    if (restoreFocus) {
      window.setTimeout(() => slideFrameRef.current?.focus({ preventScroll: true }), 0);
    }
  }, []);

  const commitDirectTextEditor = useCallback(
    (restoreFocus = true): Promise<boolean> | undefined => {
      const activeEditor = directTextEditorStateRef.current;
      const editorElement = directTextEditorRef.current;
      const session = editor;
      if (activeEditor === null || editorElement === null || session === null) return;
      if (directTextCommitPromiseRef.current !== null) return directTextCommitPromiseRef.current;

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
        setMessage("Text unchanged");
        return Promise.resolve(true);
      }

      const commit = (async () => {
        busyRef.current = true;
        setOperationError("");
        try {
          const result = await session.applyAll(commands);
          syncFromEditor(session, currentIndex);
          dirtyRef.current = session.history.undoDepth !== cleanUndoDepthRef.current;
          setMessage(commandMessage("Text updated", result.warnings));
          closeDirectTextEditor(restoreFocus);
          return true;
        } catch (error) {
          setHistory(session.history);
          setOperationError(error instanceof Error ? error.message : String(error));
          if (isPptxEditorError(error) && error.code === "render-failed") {
            dirtyRef.current = session.history.undoDepth !== cleanUndoDepthRef.current;
            setMessage("Text updated; slide preview could not refresh");
            closeDirectTextEditor(restoreFocus);
            return true;
          }
          return false;
        } finally {
          busyRef.current = false;
          directTextCommitPromiseRef.current = null;
        }
      })();
      directTextCommitPromiseRef.current = commit;
      return commit;
    },
    [closeDirectTextEditor, currentIndex, editor, syncFromEditor],
  );

  const startDirectTextEditor = useCallback(
    (shape: PptxEditorShapeInfo) => {
      if (
        busyRef.current ||
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
      editor.selectShape(shape.handle);
      setSelectedShapeKey(shapeKey(shape));
      setDraftBounds(null);
      const nextEditor = {
        shapeKey: shapeKey(shape),
        bounds: shape.bounds,
        paragraphs,
      };
      directTextEditorStateRef.current = nextEditor;
      setDirectTextEditor(nextEditor);
      setMessage("Editing text");
    },
    [editor],
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
        setMessage("Text edit canceled");
        return;
      }
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        void commitDirectTextEditor();
      }
    },
    [closeDirectTextEditor, commitDirectTextEditor],
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
    (properties: TextRunProperties) => {
      if (selectedRun === undefined) return;
      return applyCommand(
        { kind: "setTextRunProperties", handle: selectedRun.handle, properties },
        "Text style updated",
      );
    },
    [applyCommand, selectedRun],
  );

  const handleClearTextProperties = useCallback(() => {
    if (selectedRun === undefined) return;
    const properties: ClearTextRunProperties = [
      "bold",
      "italic",
      "underline",
      "fontSize",
      "color",
      "typeface",
    ];
    return applyCommand(
      { kind: "clearTextRunProperties", handle: selectedRun.handle, properties },
      "Text style cleared",
    );
  }, [applyCommand, selectedRun]);

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
    );
  }, [currentIndex, currentSlide?.handle, runEditorOperation]);

  const waitForDirectTextCommit = useCallback(async () => {
    const commit = directTextCommitPromiseRef.current;
    return commit === null || (await commit);
  }, []);

  const confirmDiscardChanges = useCallback(async () => {
    if (!(await waitForDirectTextCommit())) return false;
    return (
      !dirtyRef.current ||
      window.confirm("Discard your unsaved changes and open another version of the presentation?")
    );
  }, [waitForDirectTextCommit]);

  const handleOpenPptx = useCallback(async () => {
    if (await confirmDiscardChanges()) onOpenPptx();
  }, [confirmDiscardChanges, onOpenPptx]);

  const handleOpenSample = useCallback(async () => {
    if (await confirmDiscardChanges()) onOpenSample();
  }, [confirmDiscardChanges, onOpenSample]);

  const handleAddFonts = useCallback(async () => {
    if (await confirmDiscardChanges()) onAddFonts();
  }, [confirmDiscardChanges, onAddFonts]);

  useEffect(() => {
    const confirmMessage = "Discard your unsaved changes and leave the editor?";
    const hasUnsavedChanges = () => dirtyRef.current || directTextEditorStateRef.current !== null;
    const handleBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!hasUnsavedChanges()) return;
      event.preventDefault();
      event.returnValue = "";
    };
    const handleLinkClick = (event: MouseEvent) => {
      if (!hasUnsavedChanges()) return;
      const target = event.target;
      if (!(target instanceof Element)) return;
      const link = target.closest("a");
      if (link === null || link.target === "_blank" || link.hasAttribute("download")) return;
      if (window.confirm(confirmMessage)) return;
      event.preventDefault();
      event.stopPropagation();
    };

    window.addEventListener("beforeunload", handleBeforeUnload);
    document.addEventListener("click", handleLinkClick, true);
    return () => {
      window.removeEventListener("beforeunload", handleBeforeUnload);
      document.removeEventListener("click", handleLinkClick, true);
    };
  }, []);

  const handleSelectSlide = useCallback(
    async (index: number) => {
      if (await waitForDirectTextCommit()) setCurrentIndex(index);
    },
    [waitForDirectTextCommit],
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

  const handleSlideDrop = useCallback(
    (fromIndex: number, target: SlideDropTarget) => {
      const toIndex = slideDropIndex(fromIndex, target);
      setDraggedSlideIndex(null);
      setSlideDropTarget(null);
      void handleMoveSlide(fromIndex, toIndex);
    },
    [handleMoveSlide],
  );

  const clearSlideSortDrag = useCallback((pointerId?: number) => {
    const drag = slideSortDragRef.current;
    if (drag === null || (pointerId !== undefined && drag.pointerId !== pointerId)) return;
    slideSortDragRef.current = null;
    setDraggedSlideIndex(null);
    setSlideDropTarget(null);
  }, []);

  useEffect(() => {
    const handleBlur = () => clearSlideSortDrag();
    window.addEventListener("blur", handleBlur);
    return () => {
      window.removeEventListener("blur", handleBlur);
      slideSortDragRef.current = null;
    };
  }, [clearSlideSortDrag]);

  const handleSlideKeyDown = useCallback(
    (index: number, event: React.KeyboardEvent<HTMLButtonElement>) => {
      if (!event.altKey || (event.key !== "ArrowUp" && event.key !== "ArrowDown")) return;
      const toIndex = index + (event.key === "ArrowUp" ? -1 : 1);
      if (toIndex < 0 || toIndex >= slides.length) return;
      event.preventDefault();
      void handleMoveSlide(index, toIndex);
    },
    [handleMoveSlide, slides.length],
  );

  const handleOpenImageInput = useCallback(async () => {
    if (await waitForDirectTextCommit()) imageInputRef.current?.click();
  }, [waitForDirectTextCommit]);

  const handleImageReplacement = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (file === undefined || selectedShape?.handle === undefined) return;
      const replacement = selectedShape.editableImageReplacement;
      if (replacement === undefined) return;
      if (file.size > MAX_IMAGE_REPLACEMENT_BYTES) {
        setOperationError("Replacement image must be 5 MB or smaller.");
        return;
      }
      if (file.type !== "" && file.type !== replacement.contentType) {
        setOperationError(`Replacement image must use ${replacement.contentType}.`);
        return;
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      const detectedContentType = detectImageContentType(bytes);
      if (detectedContentType !== replacement.contentType) {
        setOperationError(`Replacement image must use ${replacement.contentType}.`);
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
    [applyCommand, selectedShape],
  );

  const handleDownload = useCallback(async () => {
    if (!(await waitForDirectTextCommit())) return;
    if (busyRef.current) return;
    busyRef.current = true;
    setBusy(true);
    setOperationError("");
    try {
      const saved = editor.save();
      setHistory(saved.history);
      cleanUndoDepthRef.current = saved.history.undoDepth;
      const href = URL.createObjectURL(
        new Blob([uint8ArrayToArrayBuffer(saved.pptx)], {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }),
      );
      const link = document.createElement("a");
      link.href = href;
      link.download = downloadFileName(downloadName, fileName);
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(href), 0);
      dirtyRef.current = false;
      setMessage("PPTX downloaded");
    } catch (error) {
      setOperationError(error instanceof Error ? error.message : String(error));
    } finally {
      busyRef.current = false;
      setBusy(false);
    }
  }, [downloadName, editor, fileName, waitForDirectTextCommit]);

  if (currentSlide === undefined) {
    return (
      <div className="loading" data-testid="editor-status">
        <div className="loading-mark" aria-hidden="true" />
        <p>{message}</p>
      </div>
    );
  }

  return (
    <section className="editor-workspace" aria-label="PPTX editor" data-testid="editor-workspace">
      <div className="editor-topbar">
        <div className="editor-file">
          <label className="editor-file-name">
            <span className="visually-hidden">Presentation file name</span>
            <input
              aria-label="Presentation file name"
              disabled={busy}
              maxLength={120}
              spellCheck={false}
              style={{
                width: `${Math.min(Math.max(downloadName.length + 1, 8), 30).toString()}ch`,
              }}
              value={downloadName}
              onBlur={() => {
                if (downloadName.trim() === "") setDownloadName(fileStem(fileName));
              }}
              onChange={(event) => setDownloadName(event.target.value)}
            />
            <span aria-hidden="true">.pptx</span>
          </label>
        </div>
        <div className="editor-status" data-testid="editor-status" role="status">
          {message === "" ? null : <span>{message}</span>}
          {busy ? <span>Working...</span> : null}
        </div>
        <div className="editor-file-actions">
          <button disabled={busy} type="button" onClick={() => void handleOpenPptx()}>
            Open PPTX
          </button>
          <button disabled={busy} type="button" onClick={() => void handleOpenSample()}>
            Open sample
          </button>
          <button disabled={busy} type="button" onClick={() => void handleAddFonts()}>
            Add fonts{fontFileCount > 0 ? ` (${fontFileCount.toString()})` : ""}
          </button>
          <button className="primary-action" disabled={busy} type="button" onClick={handleDownload}>
            Download PPTX
          </button>
        </div>
      </div>

      <div className="editor-commandbar" aria-label="Editing history">
        <button disabled={busy || !history.canUndo} type="button" onClick={handleUndo}>
          Undo
        </button>
        <button disabled={busy || !history.canRedo} type="button" onClick={handleRedo}>
          Redo
        </button>
      </div>

      <div className="editor-shell">
        <aside className="editor-thumbnails" aria-label="Slides">
          {slides.map((slide, index) => (
            <button
              aria-label={`Slide ${slide.slideNumber.toString()}, position ${(index + 1).toString()} of ${slides.length.toString()}. Drag to reorder, or press Alt with Arrow Up or Arrow Down.`}
              className={[
                "editor-thumbnail",
                index === currentIndex ? "active" : "",
                index === draggedSlideIndex ? "dragging" : "",
                slideDropTarget?.index === index ? `drop-${slideDropTarget.edge}` : "",
              ]
                .filter(Boolean)
                .join(" ")}
              data-slide-index={index}
              data-testid="editor-thumbnail"
              aria-keyshortcuts="Alt+ArrowUp Alt+ArrowDown"
              key={
                slide.handle === undefined ? `slide-${index.toString()}` : handleKey(slide.handle)
              }
              type="button"
              disabled={busy}
              onClick={(event) => {
                if (suppressSlideClickRef.current) {
                  suppressSlideClickRef.current = false;
                  event.preventDefault();
                  return;
                }
                void handleSelectSlide(index);
              }}
              onKeyDown={(event) => handleSlideKeyDown(index, event)}
              onLostPointerCapture={(event) => clearSlideSortDrag(event.pointerId)}
              onPointerCancel={(event) => clearSlideSortDrag(event.pointerId)}
              onPointerDown={(event) => {
                if (event.pointerType === "touch") return;
                if (busy || slide.handle === undefined || slideSortDragRef.current !== null) {
                  event.preventDefault();
                  return;
                }
                event.currentTarget.setPointerCapture(event.pointerId);
                suppressSlideClickRef.current = false;
                slideSortDragRef.current = {
                  fromIndex: index,
                  pointerId: event.pointerId,
                  startPoint: { x: event.clientX, y: event.clientY },
                  hasMoved: false,
                };
                setDraggedSlideIndex(index);
                setSlideDropTarget(null);
              }}
              onPointerMove={(event) => {
                let drag = slideSortDragRef.current;
                if (drag === null || drag.pointerId !== event.pointerId) return;
                if (
                  !drag.hasMoved &&
                  Math.hypot(
                    event.clientX - drag.startPoint.x,
                    event.clientY - drag.startPoint.y,
                  ) >= SLIDE_DRAG_THRESHOLD_PX
                ) {
                  drag = { ...drag, hasMoved: true };
                  slideSortDragRef.current = drag;
                }
                if (!drag.hasMoved) return;
                const target = slideDropTargetAtPoint(event.clientX, event.clientY);
                setSlideDropTarget(target);
              }}
              onPointerUp={(event) => {
                const drag = slideSortDragRef.current;
                if (drag === null || drag.pointerId !== event.pointerId) return;
                const target = drag.hasMoved
                  ? slideDropTargetAtPoint(event.clientX, event.clientY)
                  : null;
                const changesPosition =
                  target !== null && slideDropIndex(drag.fromIndex, target) !== drag.fromIndex;
                suppressSlideClickRef.current =
                  drag.hasMoved && (target === null || changesPosition);
                if (suppressSlideClickRef.current) {
                  window.setTimeout(() => {
                    suppressSlideClickRef.current = false;
                  }, 0);
                }
                clearSlideSortDrag(event.pointerId);
                if (target !== null && changesPosition) handleSlideDrop(drag.fromIndex, target);
              }}
            >
              <span className="editor-thumbnail-label">
                <span>Slide {slide.slideNumber}</span>
              </span>
              <span dangerouslySetInnerHTML={{ __html: slide.svg }} />
            </button>
          ))}
        </aside>

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

        <aside className="editor-panel" aria-label="Editing controls">
          <div className="panel-group">
            <div className="panel-title">Slide</div>
            <div className="button-row">
              <button disabled={busy} type="button" onClick={handleDuplicateSlide}>
                Duplicate
              </button>
              <button
                disabled={busy || slides.length <= 1}
                type="button"
                onClick={handleDeleteSlide}
              >
                Delete
              </button>
            </div>
          </div>

          <div className="panel-group">
            <div className="panel-title">Shape</div>
            <button disabled={busy} type="button" onClick={handleAddTextBox}>
              Add text box
            </button>
            <button
              disabled={busy || selectedShape?.editableDelete !== true}
              type="button"
              onClick={handleDeleteShape}
            >
              Delete shape
            </button>
            <button
              data-testid="replace-image-button"
              disabled={busy || selectedShape?.editableImageReplacement === undefined}
              type="button"
              onClick={() => void handleOpenImageInput()}
            >
              Replace image
            </button>
            <input
              ref={imageInputRef}
              data-testid="image-replacement-input"
              disabled={busy || selectedShape?.editableImageReplacement === undefined}
              hidden
              type="file"
              accept={selectedShape?.editableImageReplacement?.accept}
              onChange={handleImageReplacement}
            />
          </div>

          <div className="panel-group">
            <div className="panel-title">Text</div>
            <select
              data-testid="text-run-select"
              disabled={busy || textRuns.length === 0}
              value={Math.min(selectedRunIndex, Math.max(textRuns.length - 1, 0))}
              onChange={(event) => setSelectedRunIndex(Number(event.target.value))}
            >
              {textRuns.length === 0 ? <option>No text run selected</option> : null}
              {textRuns.map((run, index) => (
                <option key={handleKey(run.handle)} value={index}>
                  {run.label}
                </option>
              ))}
            </select>
            <p className="panel-note">
              Double-click the selected text shape or press Enter to edit.
            </p>
            <div className="format-toolbar" role="group" aria-label="Text style">
              <button
                disabled={busy || selectedRun === undefined}
                type="button"
                onClick={() => handleApplyTextProperties({ bold: true })}
              >
                B
              </button>
              <button
                disabled={busy || selectedRun === undefined}
                type="button"
                onClick={() => handleApplyTextProperties({ italic: true })}
              >
                I
              </button>
              <button
                disabled={busy || selectedRun === undefined}
                type="button"
                onClick={() => handleApplyTextProperties({ underline: true })}
              >
                U
              </button>
              <input
                aria-label="Text color"
                disabled={busy || selectedRun === undefined}
                type="color"
                value={color}
                onChange={(event) => {
                  setColor(event.target.value);
                  void handleApplyTextProperties({
                    color: { kind: "srgb", hex: event.target.value.slice(1).toUpperCase() },
                  });
                }}
              />
            </div>
            <div className="text-property-row">
              <input
                aria-label="Font size"
                disabled={busy || selectedRun === undefined}
                min={1}
                type="number"
                value={fontSize}
                onChange={(event) => setFontSize(event.target.value)}
              />
              <button
                disabled={busy || selectedRun === undefined || !isPositiveFiniteNumber(fontSize)}
                type="button"
                onClick={() => handleApplyTextProperties({ fontSize: pt(Number(fontSize)) })}
              >
                Size
              </button>
            </div>
            <div className="text-property-row">
              <input
                aria-label="Typeface"
                disabled={busy || selectedRun === undefined}
                placeholder="Typeface"
                value={typeface}
                onChange={(event) => setTypeface(event.target.value)}
              />
              <button
                disabled={busy || selectedRun === undefined || typeface.trim() === ""}
                type="button"
                onClick={() => handleApplyTextProperties({ typeface: typeface.trim() })}
              >
                Font
              </button>
            </div>
            <button
              disabled={busy || selectedRun === undefined}
              type="button"
              onClick={handleClearTextProperties}
            >
              Clear style
            </button>
          </div>

          {operationError !== "" ? (
            <div className="error compact-error" data-testid="editor-error">
              {operationError}
            </div>
          ) : null}
        </aside>
      </div>
    </section>
  );
}

function beginDrag(
  kind: "move" | "resize",
  handle: ResizeHandle | undefined,
  shapeHandle: SourceHandle,
  event: React.PointerEvent<SVGRectElement>,
  startBounds: PptxEditorShapeBoundsPx,
  dragStateRef: React.MutableRefObject<DragState | null>,
  overlayRef: React.MutableRefObject<SVGSVGElement | null>,
) {
  const startPoint = eventPoint(overlayRef.current, event.clientX, event.clientY);
  if (startPoint === null) return;
  event.currentTarget.setPointerCapture(event.pointerId);
  dragStateRef.current = {
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

function pt(value: number): NonNullable<TextRunProperties["fontSize"]> {
  return value as NonNullable<TextRunProperties["fontSize"]>;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function isPositiveFiniteNumber(value: string): boolean {
  const number = Number(value);
  return Number.isFinite(number) && number > 0;
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

function slideDropTargetAtPoint(clientX: number, clientY: number): SlideDropTarget | null {
  const element = document.elementFromPoint(clientX, clientY);
  const thumbnail = element?.closest<HTMLElement>("[data-slide-index]");
  if (thumbnail === undefined || thumbnail === null) return null;
  const index = Number.parseInt(thumbnail.dataset.slideIndex ?? "", 10);
  if (!Number.isInteger(index)) return null;
  const bounds = thumbnail.getBoundingClientRect();
  return {
    index,
    edge: clientY < bounds.top + bounds.height / 2 ? "before" : "after",
  };
}

function slideDropIndex(fromIndex: number, target: SlideDropTarget): number {
  const insertionIndex = target.index + (target.edge === "after" ? 1 : 0);
  return insertionIndex > fromIndex ? insertionIndex - 1 : insertionIndex;
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

function fileStem(fileName: string): string {
  return fileName.replace(/\.pptx$/i, "");
}

function downloadFileName(downloadName: string, fallbackFileName: string): string {
  const normalized = downloadName
    .trim()
    .replace(/\.pptx$/i, "")
    .replace(/[/:\\\u0000-\u001f]/g, "-")
    .replace(/\.+$/g, "");
  return `${normalized === "" ? fileStem(fallbackFileName) : normalized}.pptx`;
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

function uint8ArrayToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
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
