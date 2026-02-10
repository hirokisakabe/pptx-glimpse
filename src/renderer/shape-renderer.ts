import type { ShapeElement, ConnectorElement } from "../model/shape.js";
import { renderGeometry } from "./geometry/index.js";
import { renderFillAttrs, renderOutlineAttrs } from "./fill-renderer.js";
import { renderTextBody } from "./text-renderer.js";
import { emuToPixels } from "../utils/emu.js";
import { buildTransformAttr } from "./transform.js";

export function renderShape(shape: ShapeElement): string {
  const { transform, geometry, fill, outline, textBody } = shape;
  const w = emuToPixels(transform.extentWidth);
  const h = emuToPixels(transform.extentHeight);

  const transformAttr = buildTransformAttr(transform);
  const fillResult = renderFillAttrs(fill);
  const outlineAttr = renderOutlineAttrs(outline);

  const geometrySvg = renderGeometry(geometry, w, h);

  const parts: string[] = [];
  if (fillResult.defs) {
    parts.push(fillResult.defs);
  }

  parts.push(`<g transform="${transformAttr}">`);

  if (geometrySvg) {
    // Apply fill/stroke to the geometry element
    const styledGeometry = geometrySvg.replace(/^<(\w+)/, `<$1 ${fillResult.attrs} ${outlineAttr}`);
    parts.push(styledGeometry);
  }

  if (textBody) {
    const textSvg = renderTextBody(textBody, transform);
    if (textSvg) {
      parts.push(textSvg);
    }
  }

  parts.push("</g>");
  return parts.join("");
}

export function renderConnector(connector: ConnectorElement): string {
  const { transform, outline } = connector;
  const w = emuToPixels(transform.extentWidth);
  const h = emuToPixels(transform.extentHeight);
  const transformAttr = buildTransformAttr(transform);
  const outlineAttr = renderOutlineAttrs(outline);

  return `<g transform="${transformAttr}"><line x1="0" y1="0" x2="${w}" y2="${h}" ${outlineAttr} fill="none"/></g>`;
}
