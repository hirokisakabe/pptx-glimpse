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
      <div className="slide-nav" aria-label="Slide navigation">
        <button
          aria-label="Previous slide"
          disabled={currentIndex === 0}
          onClick={() => onNavigate(currentIndex - 1)}
        >
          &larr;
        </button>
        <span data-testid="slide-counter">
          Slide {slides[currentIndex].slideNumber} / {slides.length}
        </span>
        <button
          aria-label="Next slide"
          disabled={currentIndex === slides.length - 1}
          onClick={() => onNavigate(currentIndex + 1)}
        >
          &rarr;
        </button>
      </div>
      <div
        className="slide-container"
        data-testid="slide-container"
        dangerouslySetInnerHTML={{ __html: slides[currentIndex].svg }}
      />
    </>
  );
}
