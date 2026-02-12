import type { Fill, PatternFill } from "../model/fill.js";
import type { Outline, ArrowEndpoint, ArrowSize } from "../model/line.js";
import { emuToPixels } from "../utils/emu.js";

export interface FillAttrs {
  attrs: string;
  defs: string;
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
    const id = `grad-${crypto.randomUUID()}`;

    const stops = fill.stops
      .map((s) => {
        const opacityAttr = s.color.alpha < 1 ? ` stop-opacity="${s.color.alpha}"` : "";
        return `<stop offset="${s.position * 100}%" stop-color="${s.color.hex}"${opacityAttr}/>`;
      })
      .join("");

    if (fill.gradientType === "radial") {
      const cx = (fill.centerX ?? 0.5) * 100;
      const cy = (fill.centerY ?? 0.5) * 100;
      const dx = Math.max(cx, 100 - cx);
      const dy = Math.max(cy, 100 - cy);
      const r = Math.sqrt(dx * dx + dy * dy);
      const defs = `<radialGradient id="${id}" cx="${cx}%" cy="${cy}%" r="${r}%">${stops}</radialGradient>`;
      return { attrs: `fill="url(#${id})"`, defs };
    }

    const angle = fill.angle;
    const rad = (angle * Math.PI) / 180;
    const x1 = 50 - Math.cos(rad) * 50;
    const y1 = 50 - Math.sin(rad) * 50;
    const x2 = 50 + Math.cos(rad) * 50;
    const y2 = 50 + Math.sin(rad) * 50;

    const defs = `<linearGradient id="${id}" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%">${stops}</linearGradient>`;
    return { attrs: `fill="url(#${id})"`, defs };
  }

  if (fill.type === "image") {
    const id = `imgfill-${crypto.randomUUID()}`;
    const defs = `<pattern id="${id}" patternContentUnits="objectBoundingBox" width="1" height="1"><image href="data:${fill.mimeType};base64,${fill.imageData}" width="1" height="1" preserveAspectRatio="none"/></pattern>`;
    return { attrs: `fill="url(#${id})"`, defs };
  }

  if (fill.type === "pattern") {
    return renderPatternFill(fill);
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

  if (outline.lineCap) {
    parts.push(`stroke-linecap="${outline.lineCap}"`);
  }

  if (outline.lineJoin) {
    parts.push(`stroke-linejoin="${outline.lineJoin}"`);
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

function renderPatternFill(fill: PatternFill): FillAttrs {
  const id = `patt-${crypto.randomUUID()}`;
  const fg = fill.foregroundColor.hex;
  const bg = fill.backgroundColor.hex;
  const fgAlpha = fill.foregroundColor.alpha < 1 ? ` opacity="${fill.foregroundColor.alpha}"` : "";

  const content = getPatternContent(fill.preset, fg, fgAlpha);
  if (!content) {
    const alphaAttr =
      fill.foregroundColor.alpha < 1 ? ` fill-opacity="${fill.foregroundColor.alpha}"` : "";
    return { attrs: `fill="${fg}"${alphaAttr}`, defs: "" };
  }

  const bgAlpha =
    fill.backgroundColor.alpha < 1 ? ` fill-opacity="${fill.backgroundColor.alpha}"` : "";
  const defs = `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${content.size}" height="${content.size}"><rect width="${content.size}" height="${content.size}" fill="${bg}"${bgAlpha}/>${content.svg}</pattern>`;
  return { attrs: `fill="url(#${id})"`, defs };
}

interface PatternContent {
  svg: string;
  size: number;
}

function getPatternContent(preset: string, fg: string, fgAlpha: string): PatternContent | null {
  const s = 8;
  const sw = 1;
  const line = (x1: number, y1: number, x2: number, y2: number) =>
    `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${fg}" stroke-width="${sw}"${fgAlpha}/>`;

  switch (preset) {
    case "ltHorz":
      return { svg: line(0, 4, 8, 4), size: s };
    case "ltVert":
      return { svg: line(4, 0, 4, 8), size: s };
    case "ltDnDiag":
      return { svg: line(0, 0, 8, 8), size: s };
    case "ltUpDiag":
      return { svg: line(0, 8, 8, 0), size: s };
    case "dkHorz":
      return {
        svg: line(0, 2, 8, 2) + line(0, 6, 8, 6),
        size: s,
      };
    case "dkVert":
      return {
        svg: line(2, 0, 2, 8) + line(6, 0, 6, 8),
        size: s,
      };
    case "dkDnDiag":
      return {
        svg: line(0, 0, 8, 8) + line(-4, 0, 4, 8),
        size: s,
      };
    case "dkUpDiag":
      return {
        svg: line(0, 8, 8, 0) + line(4, 8, 12, 0),
        size: s,
      };
    case "horz":
      return { svg: line(0, 4, 8, 4), size: s };
    case "vert":
      return { svg: line(4, 0, 4, 8), size: s };
    case "dnDiag":
      return { svg: line(0, 0, 8, 8), size: s };
    case "upDiag":
      return { svg: line(0, 8, 8, 0), size: s };
    case "cross":
    case "smGrid":
      return {
        svg: line(0, 4, 8, 4) + line(4, 0, 4, 8),
        size: s,
      };
    case "lgGrid":
      return {
        svg: line(0, 0, 16, 0) + line(0, 0, 0, 16),
        size: 16,
      };
    case "diagCross":
      return {
        svg: line(0, 0, 8, 8) + line(0, 8, 8, 0),
        size: s,
      };
    case "pct5":
      return {
        svg: `<rect x="0" y="0" width="1" height="1" fill="${fg}"${fgAlpha}/>`,
        size: s,
      };
    case "pct10":
      return {
        svg: `<rect x="0" y="0" width="1" height="1" fill="${fg}"${fgAlpha}/><rect x="4" y="4" width="1" height="1" fill="${fg}"${fgAlpha}/>`,
        size: s,
      };
    case "pct20":
      return {
        svg: `<rect x="0" y="0" width="2" height="2" fill="${fg}"${fgAlpha}/><rect x="4" y="4" width="2" height="2" fill="${fg}"${fgAlpha}/>`,
        size: s,
      };
    case "pct25":
      return {
        svg: `<rect x="0" y="0" width="2" height="2" fill="${fg}"${fgAlpha}/><rect x="4" y="0" width="2" height="2" fill="${fg}"${fgAlpha}/><rect x="2" y="4" width="2" height="2" fill="${fg}"${fgAlpha}/><rect x="6" y="4" width="2" height="2" fill="${fg}"${fgAlpha}/>`,
        size: s,
      };
    case "pct30":
    case "pct40":
    case "pct50":
    case "pct60":
    case "pct70":
    case "pct75":
    case "pct80":
    case "pct90": {
      const pctVal = parseInt(preset.replace("pct", ""), 10);
      const alpha = pctVal / 100;
      return {
        svg: `<rect width="${s}" height="${s}" fill="${fg}" opacity="${alpha}"${fgAlpha}/>`,
        size: s,
      };
    }
    default:
      return null;
  }
}

// --- Arrow marker rendering ---

export interface MarkerResult {
  defs: string;
  startAttr: string;
  endAttr: string;
}

const ARROW_SIZE_MAP: Record<ArrowSize, number> = {
  sm: 5,
  med: 8,
  lg: 12,
};

export function renderMarkers(outline: Outline | null): MarkerResult {
  const empty: MarkerResult = { defs: "", startAttr: "", endAttr: "" };
  if (!outline) return empty;
  if (!outline.headEnd && !outline.tailEnd) return empty;

  const color = outline.fill?.color.hex ?? "#000000";
  const alpha = outline.fill?.color.alpha ?? 1;
  const defs: string[] = [];
  let startAttr = "";
  let endAttr = "";

  if (outline.headEnd) {
    const id = `marker-${crypto.randomUUID()}`;
    const markerDef = buildMarkerDef(id, outline.headEnd, color, alpha);
    if (markerDef) {
      defs.push(markerDef);
      startAttr = `marker-start="url(#${id})"`;
    }
  }

  if (outline.tailEnd) {
    const id = `marker-${crypto.randomUUID()}`;
    const markerDef = buildMarkerDef(id, outline.tailEnd, color, alpha);
    if (markerDef) {
      defs.push(markerDef);
      endAttr = `marker-end="url(#${id})"`;
    }
  }

  return { defs: defs.join(""), startAttr, endAttr };
}

function buildMarkerDef(
  id: string,
  endpoint: ArrowEndpoint,
  color: string,
  alpha: number,
): string | null {
  const mw = ARROW_SIZE_MAP[endpoint.length];
  const mh = ARROW_SIZE_MAP[endpoint.width];
  const alphaAttr = alpha < 1 ? ` opacity="${alpha}"` : "";

  let path: string;
  let fillAttr: string;
  switch (endpoint.type) {
    case "triangle":
      path = `M 0 0 L ${mw} ${mh / 2} L 0 ${mh} Z`;
      fillAttr = `fill="${color}"`;
      break;
    case "stealth":
      path = `M 0 0 L ${mw} ${mh / 2} L 0 ${mh} L ${mw * 0.3} ${mh / 2} Z`;
      fillAttr = `fill="${color}"`;
      break;
    case "diamond":
      path = `M 0 ${mh / 2} L ${mw / 2} 0 L ${mw} ${mh / 2} L ${mw / 2} ${mh} Z`;
      fillAttr = `fill="${color}"`;
      break;
    case "oval":
      return `<marker id="${id}" markerWidth="${mw}" markerHeight="${mh}" refX="${mw / 2}" refY="${mh / 2}" orient="auto" markerUnits="userSpaceOnUse"><ellipse cx="${mw / 2}" cy="${mh / 2}" rx="${mw / 2}" ry="${mh / 2}" ${`fill="${color}"`}${alphaAttr}/></marker>`;
    case "arrow":
      path = `M 0 0 L ${mw} ${mh / 2} L 0 ${mh}`;
      fillAttr = `fill="none" stroke="${color}" stroke-width="1"`;
      break;
    default:
      return null;
  }

  return `<marker id="${id}" markerWidth="${mw}" markerHeight="${mh}" refX="${mw}" refY="${mh / 2}" orient="auto" markerUnits="userSpaceOnUse"><path d="${path}" ${fillAttr}${alphaAttr}/></marker>`;
}
