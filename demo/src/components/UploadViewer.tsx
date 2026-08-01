"use client";

import { useCallback, useEffect, useRef, useState } from "react";

import { DropZone, SAMPLE_PPTX_FILES, type SampleOpenMode, type SamplePptx } from "./DropZone";
import { EditorWorkspace } from "./EditorWorkspace";
import { type DemoMode, type LoadedPresentation, loadPresentation } from "./load-presentation";
import { SlideViewer } from "./SlideViewer";
import { ThumbnailStrip } from "./ThumbnailStrip";

interface ActivePresentation {
  readonly requestId: number;
  readonly loaded: LoadedPresentation;
}

type ReplacementState =
  | { readonly status: "idle" }
  | { readonly status: "loading" }
  | { readonly status: "error"; readonly message: string };

type ViewerState =
  | { readonly phase: "loading" }
  | { readonly phase: "error"; readonly message: string }
  | {
      readonly phase: "viewing";
      readonly presentation: ActivePresentation;
      readonly replacement: ReplacementState;
    };

export function UploadViewer() {
  const [viewerState, setViewerState] = useState<ViewerState>({ phase: "loading" });
  const [currentIndex, setCurrentIndex] = useState(0);
  const [fontFiles, setFontFiles] = useState<File[]>([]);
  const initialSampleRequested = useRef(false);
  const loadRequestIdRef = useRef(0);
  const pptxInputRef = useRef<HTMLInputElement>(null);
  const fontInputRef = useRef<HTMLInputElement>(null);

  const handleFontFiles = useCallback((files: File[]) => {
    setFontFiles(files);
  }, []);

  const handleFile = useCallback(
    async (
      file: File,
      initialMode: DemoMode = "view",
      selectedFontFiles: readonly File[] = fontFiles,
      requestId = ++loadRequestIdRef.current,
    ) => {
      setViewerState((state) =>
        state.phase === "viewing"
          ? { ...state, replacement: { status: "loading" } }
          : { phase: "loading" },
      );

      try {
        const loaded = await loadPresentation(file, initialMode, selectedFontFiles);
        if (requestId !== loadRequestIdRef.current) return;

        setCurrentIndex(0);
        setViewerState({
          phase: "viewing",
          presentation: { requestId, loaded },
          replacement: { status: "idle" },
        });
      } catch (err) {
        if (requestId !== loadRequestIdRef.current) return;
        const message = err instanceof Error ? err.message : String(err);
        setViewerState((state) =>
          state.phase === "viewing"
            ? { ...state, replacement: { status: "error", message } }
            : { phase: "error", message },
        );
      }
    },
    [fontFiles],
  );

  const handleSample = useCallback(
    async (sample: SamplePptx, initialMode: SampleOpenMode) => {
      const requestId = ++loadRequestIdRef.current;
      setViewerState((state) =>
        state.phase === "viewing"
          ? { ...state, replacement: { status: "loading" } }
          : { phase: "loading" },
      );

      try {
        const response = await fetch(sample.href);
        if (!response.ok) {
          throw new Error(`Could not load sample PPTX: ${response.status.toString()}`);
        }
        const file = new File([await response.blob()], sample.filename, {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        });
        if (requestId !== loadRequestIdRef.current) return;
        await handleFile(file, initialMode, fontFiles, requestId);
      } catch (err) {
        if (requestId !== loadRequestIdRef.current) return;
        const message = err instanceof Error ? err.message : String(err);
        setViewerState((state) =>
          state.phase === "viewing"
            ? { ...state, replacement: { status: "error", message } }
            : { phase: "error", message },
        );
      }
    },
    [fontFiles, handleFile],
  );

  useEffect(() => {
    if (initialSampleRequested.current) return;
    initialSampleRequested.current = true;
    void handleSample(SAMPLE_PPTX_FILES[0], "edit");
  }, [handleSample]);

  const handlePptxChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (file !== undefined) void handleFile(file, "edit");
    },
    [handleFile],
  );

  const handleFontChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(event.target.files ?? []);
      event.target.value = "";
      handleFontFiles(files);
      if (viewerState.phase === "viewing") {
        void handleFile(viewerState.presentation.loaded.file, "edit", files);
      }
    },
    [handleFile, handleFontFiles, viewerState],
  );

  const slides =
    viewerState.phase === "viewing" && viewerState.presentation.loaded.mode === "view"
      ? viewerState.presentation.loaded.slides
      : [];

  const handleNavigate = useCallback(
    (index: number) => {
      if (index >= 0 && index < slides.length) {
        setCurrentIndex(index);
      }
    },
    [slides.length],
  );

  switch (viewerState.phase) {
    case "loading":
      return (
        <div className="loading" data-testid="viewer-status">
          <div className="loading-mark" aria-hidden="true" />
          <p>Converting in this browser...</p>
        </div>
      );
    case "error":
      return (
        <>
          <div className="error" data-testid="viewer-error">
            {viewerState.message}
          </div>
          <DropZone
            fontFiles={fontFiles}
            onFile={handleFile}
            onFontFiles={handleFontFiles}
            onSample={handleSample}
          />
        </>
      );
    case "viewing": {
      const { presentation, replacement } = viewerState;
      const replacementProps = {
        errorMessage: replacement.status === "error" ? replacement.message : "",
        isReplacing: replacement.status === "loading",
      };

      if (presentation.loaded.mode === "edit") {
        const loaded = presentation.loaded;
        return (
          <>
            <ReplacementShell {...replacementProps}>
              <EditorWorkspace
                key={presentation.requestId}
                editor={loaded.editor}
                fileName={loaded.fileName}
                fontFileCount={fontFiles.length}
                onAddFonts={() => fontInputRef.current?.click()}
                onOpenPptx={() => pptxInputRef.current?.click()}
                onOpenSample={() => void handleSample(SAMPLE_PPTX_FILES[0], "edit")}
              />
            </ReplacementShell>
            <input
              ref={pptxInputRef}
              data-testid="pptx-input"
              type="file"
              accept=".pptx"
              hidden
              onChange={handlePptxChange}
            />
            <input
              ref={fontInputRef}
              data-testid="font-input"
              type="file"
              accept=".ttf,.otf,.ttc,font/ttf,font/otf"
              hidden
              multiple
              onChange={handleFontChange}
            />
          </>
        );
      }

      return (
        <ReplacementShell {...replacementProps}>
          <div className="viewer-toolbar">
            <div className="viewer-summary" data-testid="viewer-status">
              <span>{slides.length} slides rendered</span>
              <span>
                {presentation.loaded.fonts.length} font file
                {presentation.loaded.fonts.length === 1 ? "" : "s"} provided for this render
              </span>
            </div>
            <div className="mode-switch" role="group" aria-label="Demo mode">
              <button aria-pressed="true" type="button">
                View
              </button>
              <button
                data-testid="open-editor"
                type="button"
                onClick={() => void handleFile(presentation.loaded.file, "edit")}
              >
                Edit
              </button>
            </div>
          </div>
          <SlideViewer slides={slides} currentIndex={currentIndex} onNavigate={handleNavigate} />
          <ThumbnailStrip slides={slides} currentIndex={currentIndex} onSelect={handleNavigate} />
          <DropZone
            compact
            fontFiles={fontFiles}
            onFile={handleFile}
            onFontFiles={handleFontFiles}
            onSample={handleSample}
          />
        </ReplacementShell>
      );
    }
  }
}

function ReplacementShell({
  children,
  errorMessage,
  isReplacing,
}: {
  readonly children: React.ReactNode;
  readonly errorMessage: string;
  readonly isReplacing: boolean;
}) {
  return (
    <div className="editor-replacement-shell" aria-busy={isReplacing}>
      <div inert={isReplacing ? true : undefined}>{children}</div>
      {errorMessage === "" ? null : (
        <div className="replacement-error" role="alert">
          {errorMessage}
        </div>
      )}
      {isReplacing ? (
        <div className="replacement-loading" data-testid="replacement-loading">
          <div className="loading-mark" aria-hidden="true" />
          <span>Opening presentation...</span>
        </div>
      ) : null}
    </div>
  );
}
