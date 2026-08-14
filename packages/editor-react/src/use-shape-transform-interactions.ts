import {
  type EditorCommand,
  type PptxEditorSession,
  type PptxEditorShapeBoundsPx,
  type PptxEditorShapeInfo,
  type SourceHandle,
} from "pptx-glimpse";
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";

import type { ApplyEditorCommand } from "./editor-interaction-types.js";
import { handleKey } from "./editor-interaction-utils.js";
import type { PptxEditorController } from "./pptx-editor-controller.js";

const EMU_PER_PIXEL = 9525;
const MIN_SHAPE_SIZE = 8;

type ShapeTransformCommand = Extract<EditorCommand, { readonly kind: "setShapeTransform" }>;
type ResizeHandle = "nw" | "ne" | "sw" | "se";

interface Point {
  readonly x: number;
  readonly y: number;
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

interface UseShapeTransformInteractionsOptions {
  readonly applyCommand: ApplyEditorCommand;
  readonly controller: PptxEditorController<PptxEditorSession>;
  readonly currentIndex: number;
  readonly directTextEditing: boolean;
  readonly shapeOptions: readonly PptxEditorShapeInfo[];
  readonly selectedShapeKey: string | null;
  readonly slideFrameRef: React.RefObject<HTMLDivElement | null>;
}

export function useShapeTransformInteractions({
  applyCommand,
  controller,
  currentIndex,
  directTextEditing,
  shapeOptions,
  selectedShapeKey,
  slideFrameRef,
}: UseShapeTransformInteractionsOptions) {
  const [draftBounds, setDraftBounds] = useState<PptxEditorShapeBoundsPx | null>(null);
  const overlayRef = useRef<SVGSVGElement | null>(null);
  const dragStateRef = useRef<DragState | null>(null);
  const committedInteractionScopeRef = useRef<object | null>(null);

  const selectedShape = useMemo(() => {
    if (selectedShapeKey === null) return null;
    const shape = shapeOptions.find((candidate) => shapeKey(candidate) === selectedShapeKey);
    if (shape === undefined) return null;
    return draftBounds === null ? shape : { ...shape, bounds: draftBounds };
  }, [draftBounds, selectedShapeKey, shapeOptions]);

  const cancelGesture = useCallback(() => {
    dragStateRef.current = null;
    setDraftBounds(null);
  }, []);

  useEffect(() => {
    setDraftBounds(null);
  }, [currentIndex, selectedShapeKey]);

  useLayoutEffect(() => {
    committedInteractionScopeRef.current = controller;
    cancelGesture();
    return () => {
      if (committedInteractionScopeRef.current === controller) {
        committedInteractionScopeRef.current = null;
      }
      dragStateRef.current = null;
    };
  }, [cancelGesture, controller]);

  const handleSelectShape = useCallback(
    (shape: PptxEditorShapeInfo, event?: React.PointerEvent<SVGRectElement>) => {
      if (controller.getSnapshot().busy || directTextEditing || shape.handle === undefined) return;
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
    [controller, directTextEditing, slideFrameRef],
  );

  const updateDrag = useCallback((event: PointerEvent) => {
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
  }, []);

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
        if (sameBounds(dragState.startBounds, nextBounds)) return;
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
    [applyCommand],
  );

  useEffect(() => {
    const move = (event: PointerEvent) => updateDrag(event);
    const up = (event: PointerEvent) => void finishDrag(event);
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", cancelGesture);
    window.addEventListener("blur", cancelGesture);
    return () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", cancelGesture);
      window.removeEventListener("blur", cancelGesture);
    };
  }, [cancelGesture, finishDrag, updateDrag]);

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

  return {
    cancelGesture,
    handleResizeStart,
    handleSelectShape,
    overlayRef,
    selectedShape,
  };
}

export function handlePoint(bounds: PptxEditorShapeBoundsPx, handle: ResizeHandle): Point {
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  if (handle === "nw") return { x: bounds.x, y: bounds.y };
  if (handle === "ne") return { x: right, y: bounds.y };
  if (handle === "sw") return { x: bounds.x, y: bottom };
  return { x: right, y: bottom };
}

export function shapeKey(shape: PptxEditorShapeInfo): string {
  return shape.handle === undefined ? "" : handleKey(shape.handle);
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
  if (handle === "sw" || handle === "se") {
    next.height = Math.max(MIN_SHAPE_SIZE, bounds.height + dy);
  }
  return next;
}

function sameBounds(a: PptxEditorShapeBoundsPx, b: PptxEditorShapeBoundsPx): boolean {
  return (
    Math.round(a.x) === Math.round(b.x) &&
    Math.round(a.y) === Math.round(b.y) &&
    Math.round(a.width) === Math.round(b.width) &&
    Math.round(a.height) === Math.round(b.height)
  );
}

function pxToEmu(value: number): ShapeTransformCommand["offsetX"] {
  // This is the package-local constructor for the branded public EMU command field.
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion
  return Math.round(value * EMU_PER_PIXEL) as ShapeTransformCommand["offsetX"];
}
