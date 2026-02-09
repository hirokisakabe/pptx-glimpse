import type { ResolvedColor } from "../model/fill.js";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function applyColorTransforms(color: ResolvedColor, node: any): ResolvedColor {
  let { hex, alpha } = color;

  if (node.lumMod || node.lumOff) {
    hex = applyLuminance(hex, node.lumMod?.["@_val"], node.lumOff?.["@_val"]);
  }

  if (node.tint) {
    hex = applyTint(hex, Number(node.tint["@_val"]) / 100000);
  }

  if (node.shade) {
    hex = applyShade(hex, Number(node.shade["@_val"]) / 100000);
  }

  if (node.alpha) {
    alpha = Number(node.alpha["@_val"]) / 100000;
  }

  return { hex, alpha };
}

function applyLuminance(hex: string, lumModVal?: string, lumOffVal?: string): string {
  const { h, s, l } = hexToHsl(hex);
  const lumMod = lumModVal ? Number(lumModVal) / 100000 : 1;
  const lumOff = lumOffVal ? Number(lumOffVal) / 100000 : 0;
  const newL = Math.min(1, Math.max(0, l * lumMod + lumOff));
  return hslToHex(h, s, newL);
}

function applyTint(hex: string, tintAmount: number): string {
  const { r, g, b } = hexToRgb(hex);
  const newR = Math.round(r + (255 - r) * tintAmount);
  const newG = Math.round(g + (255 - g) * tintAmount);
  const newB = Math.round(b + (255 - b) * tintAmount);
  return rgbToHex(newR, newG, newB);
}

function applyShade(hex: string, shadeAmount: number): string {
  const { r, g, b } = hexToRgb(hex);
  const newR = Math.round(r * shadeAmount);
  const newG = Math.round(g * shadeAmount);
  const newB = Math.round(b * shadeAmount);
  return rgbToHex(newR, newG, newB);
}

function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const h = hex.replace("#", "");
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16),
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (n: number) => Math.min(255, Math.max(0, n)).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function hexToHsl(hex: string): { h: number; s: number; l: number } {
  const { r: r255, g: g255, b: b255 } = hexToRgb(hex);
  const r = r255 / 255;
  const g = g255 / 255;
  const b = b255 / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;

  if (max === min) {
    return { h: 0, s: 0, l };
  }

  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

  let h: number;
  if (max === r) {
    h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  } else if (max === g) {
    h = ((b - r) / d + 2) / 6;
  } else {
    h = ((r - g) / d + 4) / 6;
  }

  return { h, s, l };
}

function hslToHex(h: number, s: number, l: number): string {
  if (s === 0) {
    const v = Math.round(l * 255);
    return rgbToHex(v, v, v);
  }

  const hue2rgb = (p: number, q: number, t: number): number => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;

  const r = Math.round(hue2rgb(p, q, h + 1 / 3) * 255);
  const g = Math.round(hue2rgb(p, q, h) * 255);
  const b = Math.round(hue2rgb(p, q, h - 1 / 3) * 255);

  return rgbToHex(r, g, b);
}
