"use client";

import { useEffect } from "react";

interface Slide {
  slideNumber: number;
  svg: string;
}

export function SlideViewer({
  slides,
  currentIndex,
  onNavigate,
}: {
  slides: Slide[];
  currentIndex: number;
  onNavigate: (index: number) => void;
}) {
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "ArrowLeft") onNavigate(currentIndex - 1);
      if (e.key === "ArrowRight") onNavigate(currentIndex + 1);
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [currentIndex, onNavigate]);

  return (
    <>
      <div className="slide-nav">
        <button disabled={currentIndex === 0} onClick={() => onNavigate(currentIndex - 1)}>
          &laquo; Prev
        </button>
        <span>
          Slide {currentIndex + 1} / {slides.length}
        </span>
        <button
          disabled={currentIndex === slides.length - 1}
          onClick={() => onNavigate(currentIndex + 1)}
        >
          Next &raquo;
        </button>
      </div>
      <div
        className="slide-container"
        dangerouslySetInnerHTML={{ __html: slides[currentIndex].svg }}
      />
    </>
  );
}
