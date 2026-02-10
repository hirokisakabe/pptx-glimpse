import type { Slide } from "../model/slide.js";
import type { SlideSize } from "../model/presentation.js";
import type { SlideElement, GroupElement } from "../model/shape.js";
import { emuToPixels } from "../utils/emu.js";
import { renderShape, renderConnector } from "./shape-renderer.js";
import { renderImage } from "./image-renderer.js";
import { renderChart } from "./chart-renderer.js";
import { renderFillAttrs, resetDefsCounter } from "./fill-renderer.js";

export function renderSlideToSvg(slide: Slide, slideSize: SlideSize): string {
  resetDefsCounter();

  const width = emuToPixels(slideSize.width);
  const height = emuToPixels(slideSize.height);

  const parts: string[] = [];
  const defs: string[] = [];

  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">`,
  );

  // Background
  if (slide.background?.fill) {
    const fillResult = renderFillAttrs(slide.background.fill);
    if (fillResult.defs) defs.push(fillResult.defs);
    parts.push(`<rect width="${width}" height="${height}" ${fillResult.attrs}/>`);
  } else {
    parts.push(`<rect width="${width}" height="${height}" fill="#FFFFFF"/>`);
  }

  // Elements
  for (const element of slide.elements) {
    const rendered = renderElement(element, defs);
    if (rendered) parts.push(rendered);
  }

  // Insert defs if any
  if (defs.length > 0) {
    parts.splice(1, 0, `<defs>${defs.join("")}</defs>`);
  }

  parts.push("</svg>");
  return parts.join("");
}

function renderElement(element: SlideElement, defs: string[]): string | null {
  switch (element.type) {
    case "shape": {
      const result = renderShape(element);
      // Extract defs from result if present
      extractDefs(result, defs);
      return removeDefs(result);
    }
    case "image":
      return renderImage(element);
    case "connector":
      return renderConnector(element);
    case "group":
      return renderGroup(element, defs);
    case "chart":
      return renderChart(element);
    default:
      return null;
  }
}

function renderGroup(group: GroupElement, defs: string[]): string {
  const x = emuToPixels(group.transform.offsetX);
  const y = emuToPixels(group.transform.offsetY);
  const w = emuToPixels(group.transform.extentWidth);
  const h = emuToPixels(group.transform.extentHeight);
  const chW = emuToPixels(group.childTransform.extentWidth);
  const chH = emuToPixels(group.childTransform.extentHeight);
  const chX = emuToPixels(group.childTransform.offsetX);
  const chY = emuToPixels(group.childTransform.offsetY);

  const scaleX = chW !== 0 ? w / chW : 1;
  const scaleY = chH !== 0 ? h / chH : 1;

  const parts: string[] = [];
  parts.push(
    `<g transform="translate(${x}, ${y}) scale(${scaleX}, ${scaleY}) translate(${-chX}, ${-chY})">`,
  );

  for (const child of group.children) {
    const rendered = renderElement(child, defs);
    if (rendered) parts.push(rendered);
  }

  parts.push("</g>");
  return parts.join("");
}

function extractDefs(svgFragment: string, defs: string[]): void {
  const defsMatch = svgFragment.match(/<linearGradient[^]*?<\/linearGradient>/g);
  if (defsMatch) {
    defs.push(...defsMatch);
  }
}

function removeDefs(svgFragment: string): string {
  return svgFragment.replace(/<linearGradient[^]*?<\/linearGradient>/g, "");
}
