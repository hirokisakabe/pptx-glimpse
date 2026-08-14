import type { PptxEditorSlideSvg, SourceHandle } from "pptx-glimpse";

export function handleKey(handle: SourceHandle): string {
  return [
    handle.partPath ?? "",
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot === undefined ? "" : String(handle.orderingSlot),
  ].join("\u0000");
}

export function findSlideIndexByHandle(
  slides: readonly PptxEditorSlideSvg[],
  handle: SourceHandle | undefined,
  fallback: number,
): number {
  if (handle === undefined) return fallback;
  const index = slides.findIndex(
    (slide) => slide.handle !== undefined && handleKey(slide.handle) === handleKey(handle),
  );
  return index === -1 ? fallback : index;
}

export function viewBoxFromSvg(svg: string): string {
  const fallback = svg.match(/\sviewBox="([^"]+)"/)?.[1] ?? "0 0 960 540";
  if (typeof DOMParser === "undefined") return fallback;
  const parsed = new DOMParser().parseFromString(svg, "image/svg+xml");
  const root = parsed.documentElement;
  const viewBox = root.getAttribute("viewBox");
  if (viewBox !== null && viewBox.trim() !== "") return viewBox;
  const width = parsePositiveSvgLength(root.getAttribute("width"));
  const height = parsePositiveSvgLength(root.getAttribute("height"));
  return width === undefined || height === undefined ? fallback : `0 0 ${width} ${height}`;
}

function parsePositiveSvgLength(value: string | null): number | undefined {
  if (value === null) return undefined;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}
