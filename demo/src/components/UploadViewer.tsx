"use client";

import { useCallback, useState } from "react";
import { convertPptxToSvg, type FontBuffer } from "pptx-glimpse";

import { DropZone, type SamplePptx } from "./DropZone";
import { SlideViewer } from "./SlideViewer";
import { ThumbnailStrip } from "./ThumbnailStrip";

interface Slide {
  slideNumber: number;
  svg: string;
}

type Phase = "upload" | "loading" | "viewing" | "error";

export function UploadViewer() {
  const [phase, setPhase] = useState<Phase>("upload");
  const [slides, setSlides] = useState<Slide[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [errorMessage, setErrorMessage] = useState("");
  const [fontFiles, setFontFiles] = useState<File[]>([]);
  const [renderedFontCount, setRenderedFontCount] = useState(0);

  const handleFontFiles = useCallback((files: File[]) => {
    setFontFiles(files);
  }, []);

  const handleFile = useCallback(
    async (file: File) => {
      setPhase("loading");
      setErrorMessage("");

      try {
        const [pptxBytes, fonts] = await Promise.all([
          file.arrayBuffer(),
          readFontBuffers(fontFiles),
        ]);
        const report = await convertPptxToSvg(new Uint8Array(pptxBytes), {
          fonts,
          skipSystemFonts: true,
        });

        if (report.slides.length === 0) {
          throw new Error("No slides found in the selected file");
        }

        setSlides([...report.slides]);
        setRenderedFontCount(fonts.length);
        setCurrentIndex(0);
        setPhase("viewing");
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setPhase("error");
      }
    },
    [fontFiles],
  );

  const handleSample = useCallback(
    async (sample: SamplePptx) => {
      setPhase("loading");
      setErrorMessage("");

      try {
        const response = await fetch(sample.href);
        if (!response.ok) {
          throw new Error(`Could not load sample PPTX: ${response.status.toString()}`);
        }
        const file = new File([await response.blob()], sample.filename, {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        });
        await handleFile(file);
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setPhase("error");
      }
    },
    [handleFile],
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
    return (
      <>
        <div className="viewer-summary" data-testid="viewer-status">
          <span>{slides.length} slides rendered</span>
          <span>
            {renderedFontCount} font file{renderedFontCount === 1 ? "" : "s"} provided for this
            render
          </span>
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
