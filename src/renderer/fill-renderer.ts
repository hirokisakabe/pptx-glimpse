import type { Fill } from "../model/fill.js";
import type { Outline } from "../model/line.js";
import { emuToPixels } from "../utils/emu.js";

export interface FillAttrs {
  attrs: string;
  defs: string;
}

let defsCounter = 0;

export function resetDefsCounter(): void {
  defsCounter = 0;
}

export function renderFillAttrs(fill: Fill | null): FillAttrs {
  if (!fill || fill.type === "none") {
    return { attrs: `fill="none"`, defs: "" };
  }

  if (fill.type === "solid") {
    const alphaAttr = fill.color.alpha < 1 ? ` fill-opacity="${fill.color.alpha}"` : "";
    return { attrs: `fill="${fill.color.hex}"${alphaAttr}`, defs: "" };
  }

  if (fill.type === "gradient") {
    const id = `grad-${defsCounter++}`;
    const angle = fill.angle;
    const rad = (angle * Math.PI) / 180;
    const x1 = 50 - Math.cos(rad) * 50;
    const y1 = 50 - Math.sin(rad) * 50;
    const x2 = 50 + Math.cos(rad) * 50;
    const y2 = 50 + Math.sin(rad) * 50;

    const stops = fill.stops
      .map((s) => {
        const opacityAttr = s.color.alpha < 1 ? ` stop-opacity="${s.color.alpha}"` : "";
        return `<stop offset="${s.position * 100}%" stop-color="${s.color.hex}"${opacityAttr}/>`;
      })
      .join("");

    const defs = `<linearGradient id="${id}" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%">${stops}</linearGradient>`;
    return { attrs: `fill="url(#${id})"`, defs };
  }

  return { attrs: `fill="none"`, defs: "" };
}

export function renderOutlineAttrs(outline: Outline | null): string {
  if (!outline) return `stroke="none"`;

  const widthPx = emuToPixels(outline.width);
  const parts: string[] = [`stroke-width="${widthPx}"`];

  if (outline.fill) {
    parts.push(`stroke="${outline.fill.color.hex}"`);
    if (outline.fill.color.alpha < 1) {
      parts.push(`stroke-opacity="${outline.fill.color.alpha}"`);
    }
  } else {
    parts.push(`stroke="none"`);
  }

  if (outline.dashStyle !== "solid") {
    const dashArray = getDashArray(outline.dashStyle, widthPx);
    if (dashArray) {
      parts.push(`stroke-dasharray="${dashArray}"`);
    }
  }

  return parts.join(" ");
}

function getDashArray(style: string, w: number): string | null {
  const patterns: Record<string, number[]> = {
    dash: [4, 3],
    dot: [1, 3],
    dashDot: [4, 3, 1, 3],
    lgDash: [8, 3],
    lgDashDot: [8, 3, 1, 3],
    sysDash: [3, 1],
    sysDot: [1, 1],
  };

  const pattern = patterns[style];
  if (!pattern) return null;
  return pattern.map((v) => v * w).join(" ");
}
