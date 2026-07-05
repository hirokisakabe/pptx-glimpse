"use client";

interface Slide {
  slideNumber: number;
  svg: string;
}

export function ThumbnailStrip({
  slides,
  currentIndex,
  onSelect,
}: {
  slides: Slide[];
  currentIndex: number;
  onSelect: (index: number) => void;
}) {
  return (
    <div className="thumbnail-strip" aria-label="Rendered slides">
      {slides.map((slide, index) => (
        <button
          key={slide.slideNumber}
          className={`thumbnail${index === currentIndex ? " active" : ""}`}
          aria-label={`Slide ${slide.slideNumber}`}
          aria-current={index === currentIndex ? "true" : undefined}
          type="button"
          onClick={() => onSelect(index)}
          dangerouslySetInnerHTML={{ __html: slide.svg }}
        />
      ))}
    </div>
  );
}
