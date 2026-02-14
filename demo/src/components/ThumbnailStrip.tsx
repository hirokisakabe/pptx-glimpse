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
    <div className="thumbnail-strip">
      {slides.map((slide, index) => (
        <div
          key={slide.slideNumber}
          className={`thumbnail${index === currentIndex ? " active" : ""}`}
          onClick={() => onSelect(index)}
          dangerouslySetInnerHTML={{ __html: slide.svg }}
        />
      ))}
    </div>
  );
}
