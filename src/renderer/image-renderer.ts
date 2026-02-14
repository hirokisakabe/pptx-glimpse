import type { ImageElement } from "../model/image.js";
import type { RenderResult } from "./render-result.js";
import { renderEffects } from "./effect-renderer.js";
import { renderBlipEffects } from "./blip-effect-renderer.js";
import { emuToPixels } from "../utils/emu.js";
import { buildTransformAttr } from "./transform.js";

export function renderImage(image: ImageElement): RenderResult {
  const w = emuToPixels(image.transform.extentWidth);
  const h = emuToPixels(image.transform.extentHeight);
  const transformAttr = buildTransformAttr(image.transform);

  if (image.mimeType === "image/emf" || image.mimeType === "image/wmf") {
    return { content: renderPlaceholder(image.mimeType, w, h, transformAttr), defs: [] };
  }

  const effectResult = renderEffects(image.effects);
  const blipEffectResult = renderBlipEffects(image.blipEffects);

  const defs: string[] = [];
  if (effectResult.filterDefs) defs.push(effectResult.filterDefs);
  if (blipEffectResult.filterDefs) defs.push(blipEffectResult.filterDefs);

  const filterAttr = effectResult.filterAttr ? ` ${effectResult.filterAttr}` : "";
  const blipFilterAttr = blipEffectResult.filterAttr ? ` ${blipEffectResult.filterAttr}` : "";

  if (image.tile) {
    return renderTiled(image, w, h, transformAttr, filterAttr, blipFilterAttr, defs);
  }

  const imgTag = buildImageTag(image, w, h);

  let inner = imgTag;
  if (blipFilterAttr) inner = `<g${blipFilterAttr}>${inner}</g>`;
  const content = `<g transform="${transformAttr}"${filterAttr}>${inner}</g>`;

  return { content, defs };
}

function buildImageTag(image: ImageElement, w: number, h: number): string {
  const src = image.srcRect;
  const stretch = image.stretch;

  if (src) {
    const clipId = `crop-${crypto.randomUUID()}`;
    const scaledW = Math.round(w / (1 - src.left - src.right));
    const scaledH = Math.round(h / (1 - src.top - src.bottom));
    const imgX = Math.round(-src.left * scaledW);
    const imgY = Math.round(-src.top * scaledH);
    return (
      `<defs><clipPath id="${clipId}"><rect x="0" y="0" width="${w}" height="${h}"/></clipPath></defs>` +
      `<image clip-path="url(#${clipId})" href="data:${image.mimeType};base64,${image.imageData}" x="${imgX}" y="${imgY}" width="${scaledW}" height="${scaledH}" preserveAspectRatio="none"/>`
    );
  }

  if (stretch) {
    const imgX = Math.round(w * stretch.left);
    const imgY = Math.round(h * stretch.top);
    const imgW = Math.round(w * (1 - stretch.left - stretch.right));
    const imgH = Math.round(h * (1 - stretch.top - stretch.bottom));
    return `<image href="data:${image.mimeType};base64,${image.imageData}" x="${imgX}" y="${imgY}" width="${imgW}" height="${imgH}" preserveAspectRatio="none"/>`;
  }

  return `<image href="data:${image.mimeType};base64,${image.imageData}" width="${w}" height="${h}" preserveAspectRatio="none"/>`;
}

function renderTiled(
  image: ImageElement,
  w: number,
  h: number,
  transformAttr: string,
  filterAttr: string,
  blipFilterAttr: string,
  defs: string[],
): RenderResult {
  const t = image.tile!;
  const patternId = `tile-${crypto.randomUUID()}`;

  const tileW = Math.round(w * t.sx);
  const tileH = Math.round(h * t.sy);
  const offsetX = emuToPixels(t.tx);
  const offsetY = emuToPixels(t.ty);

  let imgTransform = "";
  if (t.flip === "x") {
    imgTransform = ` transform="translate(${tileW}, 0) scale(-1, 1)"`;
  } else if (t.flip === "y") {
    imgTransform = ` transform="translate(0, ${tileH}) scale(1, -1)"`;
  } else if (t.flip === "xy") {
    imgTransform = ` transform="translate(${tileW}, ${tileH}) scale(-1, -1)"`;
  }

  const patternDef = `<pattern id="${patternId}" patternUnits="userSpaceOnUse" x="${offsetX}" y="${offsetY}" width="${tileW}" height="${tileH}"><image href="data:${image.mimeType};base64,${image.imageData}" width="${tileW}" height="${tileH}" preserveAspectRatio="none"${imgTransform}/></pattern>`;
  defs.push(patternDef);

  let inner = `<rect width="${w}" height="${h}" fill="url(#${patternId})"/>`;
  if (blipFilterAttr) inner = `<g${blipFilterAttr}>${inner}</g>`;
  const content = `<g transform="${transformAttr}"${filterAttr}>${inner}</g>`;

  return { content, defs };
}

function renderPlaceholder(mimeType: string, w: number, h: number, transformAttr: string): string {
  const label = mimeType === "image/emf" ? "[EMF]" : "[WMF]";
  const fontSize = Math.max(8, Math.min(Math.min(w, h) / 8, 24));
  return [
    `<g transform="${transformAttr}">`,
    `<rect width="${w}" height="${h}" fill="#E0E0E0" stroke="#BDBDBD" stroke-width="1"/>`,
    `<text x="${w / 2}" y="${h / 2}" text-anchor="middle" dominant-baseline="central" font-size="${fontSize}" fill="#757575" font-family="sans-serif">${label}</text>`,
    `</g>`,
  ].join("");
}
