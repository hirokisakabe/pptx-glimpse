import type { Transform } from "../model/shape.js";
import { emuToPixels } from "../utils/emu.js";

export function buildTransformAttr(t: Transform): string {
  const x = emuToPixels(t.offsetX);
  const y = emuToPixels(t.offsetY);
  const w = emuToPixels(t.extentWidth);
  const h = emuToPixels(t.extentHeight);

  const parts: string[] = [];
  parts.push(`translate(${x}, ${y})`);

  if (t.rotation !== 0) {
    parts.push(`rotate(${t.rotation}, ${w / 2}, ${h / 2})`);
  }

  if (t.flipH || t.flipV) {
    const sx = t.flipH ? -1 : 1;
    const sy = t.flipV ? -1 : 1;
    parts.push(`translate(${t.flipH ? w : 0}, ${t.flipV ? h : 0})`);
    parts.push(`scale(${sx}, ${sy})`);
  }

  return parts.join(" ");
}
