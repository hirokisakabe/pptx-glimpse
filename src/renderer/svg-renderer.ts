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
      rendered = extractAndRemoveDefs(renderShape(element), defs);
      break;
    }
    case "image": {
      rendered = extractAndRemoveDefs(renderImage(element), defs);
      break;
    }
    case "connector": {
      rendered = extractAndRemoveDefs(renderConnector(element), defs);
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

  if (rendered && "hyperlink" in element && element.hyperlink) {
    const href = escapeXmlAttr(element.hyperlink.url);
    rendered = `<a href="${href}">${rendered}</a>`;
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

// 5 種類の defs タグを 1 つの正規表現で抽出・除去する。
// 各パターンは開始タグと終了タグが一致するため、誤マッチのリスクはない。
const DEFS_TAGS = ["linearGradient", "radialGradient", "pattern", "filter", "marker"];
const DEFS_RE = new RegExp(DEFS_TAGS.map((tag) => `<${tag}[^]*?<\\/${tag}>`).join("|"), "g");

function extractAndRemoveDefs(svgFragment: string, defs: string[]): string {
  return svgFragment.replace(DEFS_RE, (match) => {
    defs.push(match);
    return "";
  });
}

function escapeXmlAttr(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
