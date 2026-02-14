import type { Slide } from "../model/slide.js";
import type { SlideSize } from "../model/presentation.js";
import type { SlideElement, GroupElement } from "../model/shape.js";
import type { RenderResult } from "./render-result.js";
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
    const result = renderElement(element);
    if (result) {
      parts.push(result.content);
      defs.push(...result.defs);
    }
  }

  // Insert defs if any
  if (defs.length > 0) {
    parts.splice(1, 0, `<defs>${defs.join("")}</defs>`);
  }

  parts.push("</svg>");
  return parts.join("");
}

function renderElement(element: SlideElement): RenderResult | null {
  let result: RenderResult | null = null;
  switch (element.type) {
    case "shape":
      result = renderShape(element);
      break;
    case "image":
      result = renderImage(element);
      break;
    case "connector":
      result = renderConnector(element);
      break;
    case "group":
      result = renderGroup(element);
      break;
    case "chart":
      result = renderChart(element);
      break;
    case "table":
      result = renderTable(element);
      break;
    default:
      return null;
  }

  if (result && "altText" in element && element.altText) {
    result = { ...result, content: addAriaLabel(result.content, element.altText) };
  }

  if (result && "hyperlink" in element && element.hyperlink) {
    const href = escapeXmlAttr(element.hyperlink.url);
    result = { ...result, content: `<a href="${href}">${result.content}</a>` };
  }

  return result;
}

function addAriaLabel(svgFragment: string, altText: string): string {
  const escaped = altText
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return svgFragment.replace(/^<(g|image|path)\b/, `<$1 role="img" aria-label="${escaped}"`);
}

function renderGroup(group: GroupElement): RenderResult {
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
  const defs: string[] = [];
  parts.push(`<g transform="${transformParts.join(" ")}">`);

  for (const child of group.children) {
    const childResult = renderElement(child);
    if (childResult) {
      parts.push(childResult.content);
      defs.push(...childResult.defs);
    }
  }

  parts.push("</g>");
  return { content: parts.join(""), defs };
}

function escapeXmlAttr(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
