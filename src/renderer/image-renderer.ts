import type { ImageElement } from "../model/image.js";
import { emuToPixels } from "../utils/emu.js";
import { buildTransformAttr } from "./transform.js";

export function renderImage(image: ImageElement): string {
  const w = emuToPixels(image.transform.extentWidth);
  const h = emuToPixels(image.transform.extentHeight);
  const transformAttr = buildTransformAttr(image.transform);

  return `<g transform="${transformAttr}"><image href="data:${image.mimeType};base64,${image.imageData}" width="${w}" height="${h}" preserveAspectRatio="none"/></g>`;
}
