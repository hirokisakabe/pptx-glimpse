"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import type { PptxEditorSlideSvg, SourceHandle } from "pptx-glimpse";

const SLIDE_DRAG_THRESHOLD_PX = 4;

interface EditorSlideStripProps {
  readonly slides: readonly PptxEditorSlideSvg[];
  readonly currentIndex: number;
  readonly busy: boolean;
  readonly onMove: (fromIndex: number, toIndex: number) => void;
  readonly onSelect: (index: number) => void;
}

interface SlideDropTarget {
  readonly index: number;
  readonly edge: "before" | "after";
}

interface SlideSortDragState {
  readonly fromIndex: number;
  readonly pointerId: number;
  readonly startPoint: { readonly x: number; readonly y: number };
  readonly hasMoved: boolean;
}

/** Owns slide navigation plus pointer and keyboard reordering for the reusable editor UI. */
export function EditorSlideStrip({
  slides,
  currentIndex,
  busy,
  onMove,
  onSelect,
}: EditorSlideStripProps) {
  const [draggedSlideIndex, setDraggedSlideIndex] = useState<number | null>(null);
  const [slideDropTarget, setSlideDropTarget] = useState<SlideDropTarget | null>(null);
  const slideSortDragRef = useRef<SlideSortDragState | null>(null);
  const suppressSlideClickRef = useRef(false);

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
      onMove(index, toIndex);
    },
    [onMove, slides.length],
  );

  return (
    <aside className="editor-thumbnails" aria-label="Slides" data-editor-component="slide-strip">
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
          key={slide.handle === undefined ? `slide-${index.toString()}` : handleKey(slide.handle)}
          type="button"
          disabled={busy}
          onClick={(event) => {
            if (suppressSlideClickRef.current) {
              suppressSlideClickRef.current = false;
              event.preventDefault();
              return;
            }
            onSelect(index);
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
              Math.hypot(event.clientX - drag.startPoint.x, event.clientY - drag.startPoint.y) >=
                SLIDE_DRAG_THRESHOLD_PX
            ) {
              drag = { ...drag, hasMoved: true };
              slideSortDragRef.current = drag;
            }
            if (!drag.hasMoved) return;
            setSlideDropTarget(slideDropTargetAtPoint(event.clientX, event.clientY));
          }}
          onPointerUp={(event) => {
            const drag = slideSortDragRef.current;
            if (drag === null || drag.pointerId !== event.pointerId) return;
            const target = drag.hasMoved
              ? slideDropTargetAtPoint(event.clientX, event.clientY)
              : null;
            const changesPosition =
              target !== null && slideDropIndex(drag.fromIndex, target) !== drag.fromIndex;
            suppressSlideClickRef.current = drag.hasMoved && (target === null || changesPosition);
            if (suppressSlideClickRef.current) {
              window.setTimeout(() => {
                suppressSlideClickRef.current = false;
              }, 0);
            }
            clearSlideSortDrag(event.pointerId);
            if (target !== null && changesPosition) {
              onMove(drag.fromIndex, slideDropIndex(drag.fromIndex, target));
            }
          }}
        >
          <span className="editor-thumbnail-label">
            <span>Slide {slide.slideNumber}</span>
          </span>
          <span dangerouslySetInnerHTML={{ __html: slide.svg }} />
        </button>
      ))}
    </aside>
  );
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

function handleKey(handle: SourceHandle): string {
  return [
    handle.partPath ?? "",
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot === undefined ? "" : String(handle.orderingSlot),
  ].join("\u0000");
}
