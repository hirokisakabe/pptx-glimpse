"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { convertPptxToSvg, type FontBuffer } from "pptx-glimpse";

import { DropZone, SAMPLE_PPTX_FILES, type SampleOpenMode, type SamplePptx } from "./DropZone";
import { EditorWorkspace } from "./EditorWorkspace";
import { SlideViewer } from "./SlideViewer";
import { ThumbnailStrip } from "./ThumbnailStrip";

interface Slide {
  slideNumber: number;
  svg: string;
}

type Phase = "upload" | "loading" | "viewing" | "error";
type DemoMode = "view" | "edit";

export function UploadViewer() {
  const [phase, setPhase] = useState<Phase>("loading");
  const [slides, setSlides] = useState<Slide[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [errorMessage, setErrorMessage] = useState("");
  const [isReplacing, setIsReplacing] = useState(false);
  const [fontFiles, setFontFiles] = useState<File[]>([]);
  const [fontBuffers, setFontBuffers] = useState<FontBuffer[]>([]);
  const [renderedFontCount, setRenderedFontCount] = useState(0);
  const [sourceFile, setSourceFile] = useState<{
    file: File;
    fileName: string;
    bytes: Uint8Array;
  } | null>(null);
  const [mode, setMode] = useState<DemoMode>("view");
  const initialSampleRequested = useRef(false);
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
    ) => {
      const replacing = sourceFile !== null;
      if (replacing) setIsReplacing(true);
      else setPhase("loading");
      setErrorMessage("");

      try {
        const [pptxArrayBuffer, fonts] = await Promise.all([
          file.arrayBuffer(),
          readFontBuffers(selectedFontFiles),
        ]);
        const pptxBytes = new Uint8Array(pptxArrayBuffer);
        const report = await convertPptxToSvg(new Uint8Array(pptxBytes), {
          fonts,
          skipSystemFonts: true,
        });

        if (report.slides.length === 0) {
          throw new Error("No slides found in the selected file");
        }

        setSlides([...report.slides]);
        setSourceFile({ file, fileName: file.name, bytes: pptxBytes });
        setFontBuffers(fonts);
        setRenderedFontCount(fonts.length);
        setCurrentIndex(0);
        setMode(initialMode);
        setPhase("viewing");
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setPhase(replacing ? "viewing" : "error");
      } finally {
        if (replacing) setIsReplacing(false);
      }
    },
    [fontFiles, sourceFile],
  );

  const handleSample = useCallback(
    async (sample: SamplePptx, initialMode: SampleOpenMode) => {
      const replacing = sourceFile !== null;
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
        await handleFile(file, initialMode);
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setPhase(replacing ? "viewing" : "error");
        if (replacing) setIsReplacing(false);
      }
    },
    [handleFile, sourceFile],
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
      if (sourceFile !== null) void handleFile(sourceFile.file, "edit", files);
    },
    [handleFile, handleFontFiles, sourceFile],
  );

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
    if (mode === "edit" && sourceFile !== null) {
      return (
        <>
          <div className="editor-replacement-shell" aria-busy={isReplacing}>
            <EditorWorkspace
              fileName={sourceFile.fileName}
              fontFileCount={fontFiles.length}
              fonts={fontBuffers}
              pptxBytes={sourceFile.bytes}
              onAddFonts={() => fontInputRef.current?.click()}
              onOpenPptx={() => pptxInputRef.current?.click()}
              onOpenSample={() => void handleSample(SAMPLE_PPTX_FILES[0], "edit")}
            />
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
              {renderedFontCount} font file{renderedFontCount === 1 ? "" : "s"} provided for this
              render
            </span>
          </div>
          <div className="mode-switch" role="group" aria-label="Demo mode">
            <button aria-pressed="true" type="button">
              View
            </button>
            <button data-testid="open-editor" type="button" onClick={() => setMode("edit")}>
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

async function readFontBuffers(files: readonly File[]): Promise<FontBuffer[]> {
  return Promise.all(
    files.map(async (file) => ({
      name: file.name.replace(/\.(?:ttf|otf|ttc)$/i, ""),
      data: await file.arrayBuffer(),
    })),
  );
}
