import type { ImageElement } from "../model/image.js";
import { renderEffects } from "./effect-renderer.js";
import { emuToPixels } from "../utils/emu.js";
import { buildTransformAttr } from "./transform.js";

export function renderImage(image: ImageElement): string {
  const w = emuToPixels(image.transform.extentWidth);
  const h = emuToPixels(image.transform.extentHeight);
  const transformAttr = buildTransformAttr(image.transform);
  const effectResult = renderEffects(image.effects);

  const parts: string[] = [];
  if (effectResult.filterDefs) {
    parts.push(effectResult.filterDefs);
  }
  const filterAttr = effectResult.filterAttr ? ` ${effectResult.filterAttr}` : "";

  const src = image.srcRect;
  if (src) {
    const clipId = `crop-${crypto.randomUUID()}`;
    const scaledW = Math.round(w / (1 - src.left - src.right));
    const scaledH = Math.round(h / (1 - src.top - src.bottom));
    const imgX = Math.round(-src.left * scaledW);
    const imgY = Math.round(-src.top * scaledH);
    parts.push(
      `<defs><clipPath id="${clipId}"><rect x="0" y="0" width="${w}" height="${h}"/></clipPath></defs>`,
    );
    parts.push(
      `<g transform="${transformAttr}"${filterAttr}><image clip-path="url(#${clipId})" href="data:${image.mimeType};base64,${image.imageData}" x="${imgX}" y="${imgY}" width="${scaledW}" height="${scaledH}" preserveAspectRatio="none"/></g>`,
    );
  } else {
    parts.push(
      `<g transform="${transformAttr}"${filterAttr}><image href="data:${image.mimeType};base64,${image.imageData}" width="${w}" height="${h}" preserveAspectRatio="none"/></g>`,
    );
  }

  return parts.join("");
}
