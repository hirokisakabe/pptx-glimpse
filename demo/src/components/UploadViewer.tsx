"use client";

import { useCallback, useState } from "react";
import { DropZone } from "./DropZone";
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

  const handleFile = useCallback(async (file: File) => {
    setPhase("loading");

    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch("/api/convert", { method: "POST", body: formData });
      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || "Conversion failed");
      }

      setSlides(data.slides);
      setCurrentIndex(0);
      setPhase("viewing");
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err));
      setPhase("error");
    }
  }, []);

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
      <div className="loading">
        <p>Converting...</p>
      </div>
    );
  }

  if (phase === "error") {
    return (
      <>
        <div className="error">Error: {errorMessage}</div>
        <DropZone onFile={handleFile} />
      </>
    );
  }

  if (phase === "viewing") {
    return (
      <>
        <SlideViewer slides={slides} currentIndex={currentIndex} onNavigate={handleNavigate} />
        <ThumbnailStrip slides={slides} currentIndex={currentIndex} onSelect={handleNavigate} />
      </>
    );
  }

  return <DropZone onFile={handleFile} />;
}
