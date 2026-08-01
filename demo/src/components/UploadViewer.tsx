"use client";

import { useCallback, useEffect, useRef, useState } from "react";

import { DropZone, SAMPLE_PPTX_FILES, type SampleOpenMode, type SamplePptx } from "./DropZone";
import { EditorWorkspace } from "./EditorWorkspace";
import { type DemoMode, type LoadedPresentation, loadPresentation } from "./load-presentation";
import { SlideViewer } from "./SlideViewer";
import { ThumbnailStrip } from "./ThumbnailStrip";

type Phase = "upload" | "loading" | "viewing" | "error";

interface ActivePresentation {
  readonly requestId: number;
  readonly loaded: LoadedPresentation;
}

export function UploadViewer() {
  const [phase, setPhase] = useState<Phase>("loading");
  const [currentIndex, setCurrentIndex] = useState(0);
  const [errorMessage, setErrorMessage] = useState("");
  const [isReplacing, setIsReplacing] = useState(false);
  const [fontFiles, setFontFiles] = useState<File[]>([]);
  const [presentation, setPresentation] = useState<ActivePresentation | null>(null);
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
      const replacing = presentation !== null;
      if (replacing) setIsReplacing(true);
      else setPhase("loading");
      setErrorMessage("");

      try {
        const loaded = await loadPresentation(file, initialMode, selectedFontFiles);
        if (requestId !== loadRequestIdRef.current) return;

        setPresentation({ requestId, loaded });
        setCurrentIndex(0);
        setPhase("viewing");
      } catch (err) {
        if (requestId !== loadRequestIdRef.current) return;
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setPhase(replacing ? "viewing" : "error");
      } finally {
        if (replacing && requestId === loadRequestIdRef.current) setIsReplacing(false);
      }
    },
    [fontFiles, presentation],
  );

  const handleSample = useCallback(
    async (sample: SamplePptx, initialMode: SampleOpenMode) => {
      const requestId = ++loadRequestIdRef.current;
      const replacing = presentation !== null;
      if (replacing) setIsReplacing(true);
      else setPhase("loading");
      setErrorMessage("");

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
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setPhase(replacing ? "viewing" : "error");
        if (replacing) setIsReplacing(false);
      }
    },
    [fontFiles, handleFile, presentation],
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
      if (presentation !== null) void handleFile(presentation.loaded.file, "edit", files);
    },
    [handleFile, handleFontFiles, presentation],
  );

  const slides = presentation?.loaded.mode === "view" ? presentation.loaded.slides : [];

  const handleNavigate = useCallback(
    (index: number) => {
      if (index >= 0 && index < slides.length) {
        setCurrentIndex(index);
      }
    },
    [slides.length],
  );

  if (phase === "loading") {
    return (
      <div className="loading" data-testid="viewer-status">
        <div className="loading-mark" aria-hidden="true" />
        <p>Converting in this browser...</p>
      </div>
    );
  }

  if (phase === "error") {
    return (
      <>
        <div className="error" data-testid="viewer-error">
          {errorMessage}
        </div>
        <DropZone
          fontFiles={fontFiles}
          onFile={handleFile}
          onFontFiles={handleFontFiles}
          onSample={handleSample}
        />
      </>
    );
  }

  if (phase === "viewing") {
    if (presentation?.loaded.mode === "edit") {
      const loaded = presentation.loaded;
      return (
        <>
          <div className="editor-replacement-shell" aria-busy={isReplacing}>
            <div inert={isReplacing ? true : undefined}>
              <EditorWorkspace
                key={presentation.requestId}
                editor={loaded.editor}
                fileName={loaded.fileName}
                fontFileCount={fontFiles.length}
                onAddFonts={() => fontInputRef.current?.click()}
                onOpenPptx={() => pptxInputRef.current?.click()}
                onOpenSample={() => void handleSample(SAMPLE_PPTX_FILES[0], "edit")}
              />
            </div>
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
      <>
        <div className="viewer-toolbar">
          <div className="viewer-summary" data-testid="viewer-status">
            <span>{slides.length} slides rendered</span>
            <span>
              {presentation?.loaded.fonts.length ?? 0} font file
              {presentation?.loaded.fonts.length === 1 ? "" : "s"} provided for this render
            </span>
          </div>
          <div className="mode-switch" role="group" aria-label="Demo mode">
            <button aria-pressed="true" type="button">
              View
            </button>
            <button
              data-testid="open-editor"
              type="button"
              onClick={() => {
                if (presentation !== null) void handleFile(presentation.loaded.file, "edit");
              }}
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
      </>
    );
  }

  return (
    <DropZone
      fontFiles={fontFiles}
      onFile={handleFile}
      onFontFiles={handleFontFiles}
      onSample={handleSample}
    />
  );
}
