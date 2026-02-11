import { convertFile, type SlideSvg } from "./converter.js";

let slides: SlideSvg[] = [];
let currentIndex = 0;

export function setupUI(): void {
  const dropZone = document.getElementById("drop-zone")!;
  const fileInput = document.getElementById("file-input") as HTMLInputElement;
  const viewer = document.getElementById("viewer")!;
  const loading = document.getElementById("loading")!;
  const errorDiv = document.getElementById("error")!;

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("drag-over");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    const file = e.dataTransfer?.files[0];
    if (file && file.name.endsWith(".pptx")) {
      handleFile(file);
    }
  });

  fileInput.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    if (file) handleFile(file);
  });

  document.getElementById("prev-btn")!.addEventListener("click", () => navigate(-1));
  document.getElementById("next-btn")!.addEventListener("click", () => navigate(1));

  document.addEventListener("keydown", (e) => {
    if (slides.length === 0) return;
    if (e.key === "ArrowLeft") navigate(-1);
    if (e.key === "ArrowRight") navigate(1);
  });

  async function handleFile(file: File): Promise<void> {
    dropZone.hidden = true;
    viewer.hidden = true;
    loading.hidden = false;
    errorDiv.hidden = true;

    try {
      slides = await convertFile(file);
      currentIndex = 0;
      loading.hidden = true;
      viewer.hidden = false;
      renderCurrentSlide();
      renderThumbnails();
    } catch (err) {
      loading.hidden = true;
      errorDiv.hidden = false;
      errorDiv.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
      dropZone.hidden = false;
    }
  }

  function navigate(delta: number): void {
    const newIndex = currentIndex + delta;
    if (newIndex >= 0 && newIndex < slides.length) {
      currentIndex = newIndex;
      renderCurrentSlide();
      updateThumbnailSelection();
    }
  }

  function renderCurrentSlide(): void {
    const container = document.getElementById("slide-container")!;
    container.innerHTML = slides[currentIndex].svg;
    const svg = container.querySelector("svg");
    if (svg) {
      svg.removeAttribute("width");
      svg.removeAttribute("height");
      svg.style.width = "100%";
      svg.style.height = "auto";
    }
    updateNavigation();
  }

  function updateNavigation(): void {
    const prevBtn = document.getElementById("prev-btn") as HTMLButtonElement;
    const nextBtn = document.getElementById("next-btn") as HTMLButtonElement;
    const info = document.getElementById("slide-info")!;

    prevBtn.disabled = currentIndex === 0;
    nextBtn.disabled = currentIndex === slides.length - 1;
    info.textContent = `Slide ${currentIndex + 1} / ${slides.length}`;
  }

  function renderThumbnails(): void {
    const strip = document.getElementById("thumbnail-strip")!;
    strip.innerHTML = "";
    slides.forEach((slide, index) => {
      const thumb = document.createElement("div");
      thumb.className = "thumbnail" + (index === currentIndex ? " active" : "");
      thumb.innerHTML = slide.svg;
      const svg = thumb.querySelector("svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
      thumb.addEventListener("click", () => {
        currentIndex = index;
        renderCurrentSlide();
        updateThumbnailSelection();
      });
      strip.appendChild(thumb);
    });
  }

  function updateThumbnailSelection(): void {
    const thumbs = document.querySelectorAll(".thumbnail");
    thumbs.forEach((thumb, index) => {
      thumb.classList.toggle("active", index === currentIndex);
    });
  }
}
