import type { Slide } from "../model/slide.js";
import type { SlideSize } from "../model/presentation.js";
import type { SlideElement, GroupElement } from "../model/shape.js";
import { emuToPixels } from "../utils/emu.js";
import { renderShape, renderConnector } from "./shape-renderer.js";
import { renderImage } from "./image-renderer.js";
import { renderChart } from "./chart-renderer.js";
import { renderTable } from "./table-renderer.js";
import { renderFillAttrs } from "./fill-renderer.js";

// SVG 1.1 (W3C) で出力。CSS クラスは使わずインライン属性のみ使用する。
// 理由: sharp (内部で librsvg を使用) が CSS セレクタを正しく解釈しないため。

export function renderSlideToSvg(slide: Slide, slideSize: SlideSize): string {
  const width = emuToPixels(slideSize.width);
  const height = emuToPixels(slideSize.height);

  const parts: string[] = [];
  const defs: string[] = [];

  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">`,
  );

  // Background
  if (slide.background?.fill?.type === "image") {
    const bg = slide.background.fill;
    parts.push(
      `<image href="data:${bg.mimeType};base64,${bg.imageData}" width="${width}" height="${height}" preserveAspectRatio="none"/>`,
    );
  } else if (slide.background?.fill) {
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
  let rendered: string | null = null;
  switch (element.type) {
    case "shape": {
      const result = renderShape(element);
      extractDefs(result, defs);
      rendered = removeDefs(result);
      break;
    }
    case "image": {
      const imgResult = renderImage(element);
      extractDefs(imgResult, defs);
      rendered = removeDefs(imgResult);
      break;
    }
    case "connector": {
      const cxnResult = renderConnector(element);
      extractDefs(cxnResult, defs);
      rendered = removeDefs(cxnResult);
      break;
    }
    case "group":
      rendered = renderGroup(element, defs);
      break;
    case "chart":
      rendered = renderChart(element);
      break;
    case "table":
      rendered = renderTable(element, defs);
      break;
    default:
      return null;
  }

  if (rendered && "altText" in element && element.altText) {
    rendered = addAriaLabel(rendered, element.altText);
  }

  return rendered;
}

function addAriaLabel(svgFragment: string, altText: string): string {
  const escaped = altText
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return svgFragment.replace(/^<(g|image|path)\b/, `<$1 role="img" aria-label="${escaped}"`);
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

  const transformParts: string[] = [];
  transformParts.push(`translate(${x}, ${y})`);

  if (group.transform.rotation !== 0) {
    transformParts.push(`rotate(${group.transform.rotation}, ${w / 2}, ${h / 2})`);
  }

  if (group.transform.flipH || group.transform.flipV) {
    const sx = group.transform.flipH ? -1 : 1;
    const sy = group.transform.flipV ? -1 : 1;
    transformParts.push(
      `translate(${group.transform.flipH ? w : 0}, ${group.transform.flipV ? h : 0})`,
    );
    transformParts.push(`scale(${sx}, ${sy})`);
  }

  transformParts.push(`scale(${scaleX}, ${scaleY})`);
  transformParts.push(`translate(${-chX}, ${-chY})`);

  const parts: string[] = [];
  parts.push(`<g transform="${transformParts.join(" ")}">`);

  for (const child of group.children) {
    const rendered = renderElement(child, defs);
    if (rendered) parts.push(rendered);
  }

  parts.push("</g>");
  return parts.join("");
}

function extractDefs(svgFragment: string, defs: string[]): void {
  const linearGradientMatch = svgFragment.match(/<linearGradient[^]*?<\/linearGradient>/g);
  if (linearGradientMatch) {
    defs.push(...linearGradientMatch);
  }
  const radialGradientMatch = svgFragment.match(/<radialGradient[^]*?<\/radialGradient>/g);
  if (radialGradientMatch) {
    defs.push(...radialGradientMatch);
  }
  const patternMatch = svgFragment.match(/<pattern[^]*?<\/pattern>/g);
  if (patternMatch) {
    defs.push(...patternMatch);
  }
  const filterMatch = svgFragment.match(/<filter[^]*?<\/filter>/g);
  if (filterMatch) {
    defs.push(...filterMatch);
  }
  const markerMatch = svgFragment.match(/<marker[^]*?<\/marker>/g);
  if (markerMatch) {
    defs.push(...markerMatch);
  }
}

function removeDefs(svgFragment: string): string {
  return svgFragment
    .replace(/<linearGradient[^]*?<\/linearGradient>/g, "")
    .replace(/<radialGradient[^]*?<\/radialGradient>/g, "")
    .replace(/<pattern[^]*?<\/pattern>/g, "")
    .replace(/<filter[^]*?<\/filter>/g, "")
    .replace(/<marker[^]*?<\/marker>/g, "");
}
