"use client";

import {
  type EditorCommand,
  type PptxEditorSession,
  type PptxEditorShapeInfo,
  type SourceHandle,
} from "pptx-glimpse";
import { type ReactNode, useCallback, useMemo, useRef } from "react";

import type { RunEditorOperation } from "./editor-interaction-types.js";
import { viewBoxFromSvg } from "./editor-interaction-utils.js";
import { EditorLayoutPicker } from "./EditorLayoutPicker.js";
import { EditorSlideStrip } from "./EditorSlideStrip.js";
import { EditorHistoryToolbar, type EditorTextRunOption, EditorToolbar } from "./EditorToolbar.js";
import { DirectTextEditorOverlay, useDirectTextEditor } from "./use-direct-text-editor.js";
import { useHostSaveControls } from "./use-host-save-controls.js";
import { useMediaOperations } from "./use-media-operations.js";
import { usePptxEditorController } from "./use-pptx-editor-controller.js";
import {
  handlePoint,
  shapeKey,
  useShapeTransformInteractions,
} from "./use-shape-transform-interactions.js";
import { useSlideLayoutOperations } from "./use-slide-layout-operations.js";

type EditorSession = PptxEditorSession;
type TextRunProperties = Extract<
  EditorCommand,
  { readonly kind: "setTextRunProperties" }
>["properties"];
type ClearTextRunProperties = Extract<
  EditorCommand,
  { readonly kind: "clearTextRunProperties" }
>["properties"];

export interface PptxEditorProps {
  readonly session: EditorSession;
  readonly children?: (controls: PptxEditorHostControls) => ReactNode;
}

export interface PptxEditorHostControls {
  readonly busy: boolean;
  readonly dirty: boolean;
  readonly message: string;
  readonly resetScope: object;
  readonly commitPendingEdits: () => Promise<boolean>;
  readonly hasUnsavedChanges: () => boolean;
  readonly markSaved: (
    history: ReturnType<PptxEditorSession["save"]>["history"],
    message: string,
  ) => boolean;
  readonly save: () => Promise<ReturnType<PptxEditorSession["save"]> | undefined>;
  readonly setError: (error: string) => void;
}

/** Reusable editing UI. Document loading, download policy, and site navigation stay in its host. */
export function PptxEditor({ session, children }: PptxEditorProps) {
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
  } = usePptxEditorController(session);
  const slideFrameRef = useRef<HTMLDivElement | null>(null);
  const directText = useDirectTextEditor({ controller, currentIndex, session, slideFrameRef });

  const runEditorOperation = useCallback<RunEditorOperation>(
    async (operation, success, preferredIndex = currentIndex, historyAction = "mutation") => {
      if (!(await directText.waitForCurrentCommit())) return false;
      return controller.run(operation, { success, preferredIndex, historyAction });
    },
    [controller, currentIndex, directText.waitForCurrentCommit],
  );

  const applyCommand = useCallback(
    (command: EditorCommand, success: string) =>
      runEditorOperation(async (activeSession) => {
        const result = await activeSession.apply(command);
        return commandMessage(success, result.warnings);
      }, success),
    [runEditorOperation],
  );

  const shapeTransform = useShapeTransformInteractions({
    applyCommand,
    controller,
    currentIndex,
    directTextEditing: directText.isEditing,
    shapeOptions,
    selectedShapeKey,
    slideFrameRef,
  });
  const currentSlide = slides[currentIndex];

  const textRuns = useMemo<EditorTextRunOption[]>(() => {
    const sourceShapes =
      shapeTransform.selectedShape === null ? [] : [shapeTransform.selectedShape];
    return sourceShapes.flatMap((shape) =>
      (shape.textRuns ?? []).map((run, index) => ({
        label: `${shape.name ?? shape.kind} / run ${(index + 1).toString()}`,
        text: run.text,
        handle: run.handle,
      })),
    );
  }, [shapeTransform.selectedShape]);

  const slideOperations = useSlideLayoutOperations({
    commitPendingEdits: directText.commitPendingEdits,
    controller,
    currentIndex,
    currentSlide,
    runEditorOperation,
    session,
    slides,
  });
  const mediaOperations = useMediaOperations({
    applyCommand,
    commitPendingEdits: directText.commitPendingEdits,
    controller,
    selectedShape: shapeTransform.selectedShape,
  });
  const hostControls = useHostSaveControls({
    busy,
    commitPendingEdits: directText.commitPendingEdits,
    controller,
    dirty,
    directTextEditing: directText.isEditing,
    message,
  });

  const startDirectTextEdit = useCallback(
    (shape: PptxEditorShapeInfo) => {
      shapeTransform.cancelGesture();
      directText.start(shape);
    },
    [directText.start, shapeTransform.cancelGesture],
  );

  const handleSlideFrameKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (
        event.key !== "Enter" ||
        event.target !== event.currentTarget ||
        shapeTransform.selectedShape === null ||
        directText.isEditing
      ) {
        return;
      }
      event.preventDefault();
      startDirectTextEdit(shapeTransform.selectedShape);
    },
    [directText.isEditing, shapeTransform.selectedShape, startDirectTextEdit],
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
      runEditorOperation(async (activeSession) => {
        await activeSession.addTextBox(currentIndex + 1);
      }, "Text box added"),
    [currentIndex, runEditorOperation],
  );
  const handleDeleteShape = useCallback(() => {
    if (shapeTransform.selectedShape?.handle === undefined) return;
    return applyCommand(
      { kind: "deleteShape", handle: shapeTransform.selectedShape.handle },
      "Shape deleted",
    );
  }, [applyCommand, shapeTransform.selectedShape]);

  if (currentSlide === undefined) {
    return (
      <section
        className="pptx-glimpse-editor editor-workspace"
        aria-label="PPTX editor"
        data-testid="editor-workspace"
      >
        <div className="loading" data-testid="editor-status">
          <div className="loading-mark" aria-hidden="true" />
          <p>{message}</p>
        </div>
      </section>
    );
  }

  return (
    <section
      className="pptx-glimpse-editor editor-workspace"
      aria-label="PPTX editor"
      data-testid="editor-workspace"
      data-editor-component="surface"
    >
      {children?.(hostControls)}
      <EditorHistoryToolbar
        busy={busy}
        history={history}
        onRedo={() => void slideOperations.redo()}
        onUndo={() => void slideOperations.undo()}
      >
        <EditorLayoutPicker
          busy={busy}
          catalog={session.layoutCatalog}
          currentLayoutPartPath={session.document.slides[currentIndex]?.layoutPartPath}
          interactionScope={controller}
          onAdd={(layout) => {
            const layoutPartPath = layout.handle.partPath;
            return layoutPartPath === undefined
              ? Promise.resolve(false)
              : slideOperations.addSlideFromLayout(layoutPartPath);
          }}
          previewLayout={slideOperations.previewLayout}
        />
      </EditorHistoryToolbar>

      <div className="editor-shell">
        <EditorSlideStrip
          busy={busy}
          currentIndex={currentIndex}
          interactionScope={controller}
          slides={slides}
          onMove={(fromIndex, toIndex) => void slideOperations.moveSlide(fromIndex, toIndex)}
          onSelect={(index) => void slideOperations.selectSlide(index)}
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
              ref={shapeTransform.overlayRef}
              className={`editor-selection-overlay${directText.isEditing ? " editing" : ""}`}
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
                    onPointerDown={(event) => shapeTransform.handleSelectShape(shape, event)}
                    onDoubleClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      startDirectTextEdit(shape);
                    }}
                  />
                );
              })}
              {shapeTransform.selectedShape?.bounds !== undefined ? (
                <>
                  <rect
                    className="selection-box"
                    data-testid="selection-box"
                    x={shapeTransform.selectedShape.bounds.x}
                    y={shapeTransform.selectedShape.bounds.y}
                    width={shapeTransform.selectedShape.bounds.width}
                    height={shapeTransform.selectedShape.bounds.height}
                  />
                  {shapeTransform.selectedShape.editableTransform && !busy && !directText.isEditing
                    ? (["nw", "ne", "sw", "se"] as const).map((handle) => {
                        const point = handlePoint(shapeTransform.selectedShape!.bounds!, handle);
                        return (
                          <rect
                            className={`selection-handle ${handle}`}
                            data-testid={`selection-handle-${handle}`}
                            key={handle}
                            x={point.x - 4}
                            y={point.y - 4}
                            width={8}
                            height={8}
                            onPointerDown={(event) =>
                              shapeTransform.handleResizeStart(handle, event)
                            }
                          />
                        );
                      })
                    : null}
                </>
              ) : null}
            </svg>
            {directText.editor !== null ? (
              <DirectTextEditorOverlay
                editor={directText.editor}
                editorRef={directText.editorRef}
                slideSvg={currentSlide.svg}
                onBlur={directText.handleBlur}
                onCompositionEnd={directText.handleCompositionEnd}
                onCompositionStart={directText.handleCompositionStart}
                onDone={() => void directText.save()}
                onKeyDown={directText.handleKeyDown}
              />
            ) : null}
          </div>
        </div>

        <EditorToolbar
          busy={busy}
          error={operationError}
          selectionScope={controller}
          selectedShape={shapeTransform.selectedShape}
          selectedShapeKey={selectedShapeKey}
          slidesCount={slides.length}
          textRuns={textRuns}
          onAddTextBox={() => void handleAddTextBox()}
          onApplyTextProperties={(handle, properties) =>
            void handleApplyTextProperties(handle, properties)
          }
          onClearTextProperties={(handle) => void handleClearTextProperties(handle)}
          onDeleteShape={() => void handleDeleteShape()}
          onDeleteSlide={() => void slideOperations.deleteSlide()}
          onDuplicateSlide={() => void slideOperations.duplicateSlide()}
          onReplaceImage={(file) => void mediaOperations.replaceImage(file)}
        />
      </div>
    </section>
  );
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
