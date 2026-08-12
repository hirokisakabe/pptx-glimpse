"use client";

import { useCallback, useEffect, useLayoutEffect, useState } from "react";

import type { EditorSurfaceHostControls } from "./EditorSurface";

interface DemoEditorShellProps {
  readonly controls: EditorSurfaceHostControls;
  readonly fileName: string;
  readonly fontFileCount: number;
  readonly onAddFonts: () => void;
  readonly onOpenPptx: () => void;
  readonly onOpenSample: () => void;
}

/** Demo-only host policy: presentation acquisition, download naming, and site navigation guards. */
export function DemoEditorShell({
  controls,
  fileName,
  fontFileCount,
  onAddFonts,
  onOpenPptx,
  onOpenSample,
}: DemoEditorShellProps) {
  const [downloadName, setDownloadName] = useState(() => fileStem(fileName));

  useLayoutEffect(() => {
    setDownloadName(fileStem(fileName));
  }, [controls.resetScope, fileName]);

  const confirmDiscardChanges = useCallback(async () => {
    if (!(await controls.commitPendingEdits())) return false;
    return (
      !controls.hasUnsavedChanges() ||
      window.confirm("Discard your unsaved changes and open another version of the presentation?")
    );
  }, [controls.commitPendingEdits, controls.hasUnsavedChanges]);

  useEffect(() => {
    const confirmMessage = "Discard your unsaved changes and leave the editor?";
    const handleBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!controls.hasUnsavedChanges()) return;
      event.preventDefault();
      event.returnValue = "";
    };
    const handleLinkClick = (event: MouseEvent) => {
      if (!controls.hasUnsavedChanges()) return;
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
  }, [controls.hasUnsavedChanges]);

  const handleDownload = useCallback(async () => {
    const saved = await controls.save();
    if (saved === undefined) return;
    try {
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
      controls.markSaved(saved.history, "PPTX downloaded");
    } catch (error) {
      controls.setError(error instanceof Error ? error.message : String(error));
    }
  }, [controls.markSaved, controls.save, controls.setError, downloadName, fileName]);

  const runAfterDiscardConfirmation = useCallback(
    async (action: () => void) => {
      if (await confirmDiscardChanges()) action();
    },
    [confirmDiscardChanges],
  );

  return (
    <div className="editor-topbar" data-editor-component="demo-shell">
      <div className="editor-file">
        <label className="editor-file-name">
          <span className="visually-hidden">Presentation file name</span>
          <input
            aria-label="Presentation file name"
            disabled={controls.busy}
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
        {controls.message === "" ? null : <span>{controls.message}</span>}
        {controls.busy ? <span>Working...</span> : null}
      </div>
      <div className="editor-file-actions">
        <button
          disabled={controls.busy}
          type="button"
          onClick={() => void runAfterDiscardConfirmation(onOpenPptx)}
        >
          Open PPTX
        </button>
        <button
          disabled={controls.busy}
          type="button"
          onClick={() => void runAfterDiscardConfirmation(onOpenSample)}
        >
          Open sample
        </button>
        <button
          disabled={controls.busy}
          type="button"
          onClick={() => void runAfterDiscardConfirmation(onAddFonts)}
        >
          Add fonts{fontFileCount > 0 ? ` (${fontFileCount.toString()})` : ""}
        </button>
        <button
          className="primary-action"
          disabled={controls.busy}
          type="button"
          onClick={() => void handleDownload()}
        >
          Download PPTX
        </button>
      </div>
    </div>
  );
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

function uint8ArrayToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}
