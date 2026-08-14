import type { PptxEditorSession, PptxEditorSlideSvg, SourceHandle } from "pptx-glimpse";
import { useCallback } from "react";

import type { RunEditorOperation } from "./editor-interaction-types.js";
import { findSlideIndexByHandle } from "./editor-interaction-utils.js";
import type { PptxEditorController } from "./pptx-editor-controller.js";

interface UseSlideLayoutOperationsOptions {
  readonly commitPendingEdits: () => Promise<boolean>;
  readonly controller: PptxEditorController<PptxEditorSession>;
  readonly currentIndex: number;
  readonly currentSlide: PptxEditorSlideSvg | undefined;
  readonly runEditorOperation: RunEditorOperation;
  readonly session: PptxEditorSession;
  readonly slides: readonly PptxEditorSlideSvg[];
}

export function useSlideLayoutOperations({
  commitPendingEdits,
  controller,
  currentIndex,
  currentSlide,
  runEditorOperation,
  session,
  slides,
}: UseSlideLayoutOperationsOptions) {
  const duplicateSlide = useCallback(() => {
    if (currentSlide?.handle === undefined) return;
    const nextIndex = currentIndex + 1;
    return runEditorOperation(
      async (activeSession) => {
        await activeSession.apply({ kind: "duplicateSlide", handle: currentSlide.handle! });
      },
      "Slide duplicated",
      nextIndex,
    );
  }, [currentIndex, currentSlide, runEditorOperation]);

  const addSlideFromLayout = useCallback(
    (layoutPartPath: NonNullable<SourceHandle["partPath"]>) => {
      const nextIndex = slides.length;
      return runEditorOperation(
        async (activeSession) => {
          await activeSession.apply({ kind: "addEmptySlideFromLayout", layoutPartPath });
        },
        "Slide added",
        nextIndex,
      );
    },
    [runEditorOperation, slides.length],
  );

  const previewLayout = useCallback(
    (handle: SourceHandle) => session.previewLayoutCatalogTarget(handle),
    [session],
  );

  const deleteSlide = useCallback(() => {
    if (currentSlide?.handle === undefined || slides.length <= 1) return;
    const nextIndex = Math.max(0, currentIndex - 1);
    return runEditorOperation(
      async (activeSession) => {
        await activeSession.apply({ kind: "deleteSlide", handle: currentSlide.handle! });
      },
      "Slide deleted",
      nextIndex,
    );
  }, [currentIndex, currentSlide, runEditorOperation, slides.length]);

  const undo = useCallback(() => {
    const selectedSlideHandle = currentSlide?.handle;
    return runEditorOperation(
      async (activeSession) => {
        await activeSession.undo();
      },
      "Undone",
      (activeSession) =>
        findSlideIndexByHandle(activeSession.slides, selectedSlideHandle, currentIndex),
      "undo",
    );
  }, [currentIndex, currentSlide?.handle, runEditorOperation]);

  const redo = useCallback(() => {
    const selectedSlideHandle = currentSlide?.handle;
    return runEditorOperation(
      async (activeSession) => {
        await activeSession.redo();
      },
      "Redone",
      (activeSession) =>
        findSlideIndexByHandle(activeSession.slides, selectedSlideHandle, currentIndex),
      "redo",
    );
  }, [currentIndex, currentSlide?.handle, runEditorOperation]);

  const selectSlide = useCallback(
    async (index: number) => {
      if (await commitPendingEdits()) controller.selectSlide(index);
    },
    [commitPendingEdits, controller],
  );

  const moveSlide = useCallback(
    async (fromIndex: number, toIndex: number) => {
      if (fromIndex === toIndex) return;
      const slide = slides[fromIndex];
      if (slide?.handle === undefined) return;
      const slideHandle = slide.handle;
      await runEditorOperation(
        async (activeSession) => {
          await activeSession.apply({ kind: "moveSlide", handle: slideHandle, toIndex });
        },
        `Slide ${slide.slideNumber.toString()} moved to position ${(toIndex + 1).toString()} of ${slides.length.toString()}`,
        (activeSession) => findSlideIndexByHandle(activeSession.slides, slideHandle, toIndex),
      );
    },
    [runEditorOperation, slides],
  );

  return {
    addSlideFromLayout,
    deleteSlide,
    duplicateSlide,
    moveSlide,
    previewLayout,
    redo,
    selectSlide,
    undo,
  };
}
