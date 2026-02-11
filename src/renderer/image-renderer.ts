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
  parts.push(
    `<g transform="${transformAttr}"${filterAttr}><image href="data:${image.mimeType};base64,${image.imageData}" width="${w}" height="${h}" preserveAspectRatio="none"/></g>`,
  );
  return parts.join("");
}
