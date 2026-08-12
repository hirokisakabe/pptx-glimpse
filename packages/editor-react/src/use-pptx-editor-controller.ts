"use client";

import type { PptxEditorSession } from "pptx-glimpse";
import { useMemo, useSyncExternalStore } from "react";

import { PptxEditorController, type PptxEditorControllerState } from "./pptx-editor-controller.js";

export type { PreferredSlideIndex } from "./pptx-editor-controller.js";

export function usePptxEditorController(session: PptxEditorSession): PptxEditorControllerState & {
  readonly controller: PptxEditorController<PptxEditorSession>;
} {
  const controller = useMemo(() => new PptxEditorController(session), [session]);
  const state = useSyncExternalStore(
    controller.subscribe,
    controller.getSnapshot,
    controller.getSnapshot,
  );
  return useMemo(() => ({ ...state, controller }), [controller, state]);
}
