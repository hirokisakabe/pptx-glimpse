"use client";

import type { PptxEditorSession } from "pptx-glimpse";

import { DemoEditorShell } from "./DemoEditorShell";
import { EditorSurface } from "./EditorSurface";

interface EditorWorkspaceProps {
  readonly editor: PptxEditorSession;
  readonly fileName: string;
  readonly fontFileCount: number;
  readonly onAddFonts: () => void;
  readonly onOpenPptx: () => void;
  readonly onOpenSample: () => void;
}

/** Composes the reusable editor UI with the pptx-glimpse website's file-oriented shell. */
export function EditorWorkspace({
  editor,
  fileName,
  fontFileCount,
  onAddFonts,
  onOpenPptx,
  onOpenSample,
}: EditorWorkspaceProps) {
  return (
    <EditorSurface editor={editor}>
      {(controls) => (
        <DemoEditorShell
          controls={controls}
          fileName={fileName}
          fontFileCount={fontFileCount}
          onAddFonts={onAddFonts}
          onOpenPptx={onOpenPptx}
          onOpenSample={onOpenSample}
        />
      )}
    </EditorSurface>
  );
}
