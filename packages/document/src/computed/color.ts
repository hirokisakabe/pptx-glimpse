import type {
  SourceColor,
  SourceColorMap,
  SourceColorTransform,
  SourceTheme,
} from "../source/index.js";
import type { ComputedColor } from "./pptx-computed-view.js";

const DEFAULT_COLOR_MAP: Readonly<Record<string, string>> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const FALLBACK_SCHEME_COLORS: Readonly<Record<string, string>> = {
  dk1: "#000000",
  lt1: "#ffffff",
  dk2: "#44546a",
  lt2: "#e7e6e6",
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
  hlink: "#0563c1",
  folHlink: "#954f72",
};

interface ColorResolutionContext {
  readonly theme?: SourceTheme;
  readonly colorMap: Readonly<Record<string, string>>;
}

export function buildEffectiveColorMap(
  master?: SourceColorMap,
  layoutOverride?: SourceColorMap,
  slideOverride?: SourceColorMap,
): Readonly<Record<string, string>> {
  return {
    ...DEFAULT_COLOR_MAP,
    ...master?.mapping,
    ...layoutOverride?.mapping,
    ...slideOverride?.mapping,
  };
}

export function buildComputedColorScheme(
  context: ColorResolutionContext,
): Readonly<Record<string, string>> {
  const colors: Record<string, string> = { ...FALLBACK_SCHEME_COLORS };
  for (const [name, color] of Object.entries(context.theme?.colorScheme?.colors ?? {})) {
    const resolved = resolveColor(context, color);
    if (resolved !== undefined) colors[name] = resolved.hex;
  }
  return colors;
}

export function resolveColor(
  context: ColorResolutionContext,
  color: SourceColor,
  visited: ReadonlySet<string> = new Set(),
): ComputedColor | undefined {
  const hex = resolveBaseHex(context, color, visited);
  if (hex === undefined) return undefined;

  return applyColorTransforms(hex, color.transforms);
}

function resolveBaseHex(
  context: ColorResolutionContext,
  color: SourceColor,
  visited: ReadonlySet<string>,
): string | undefined {
  switch (color.kind) {
    case "srgb":
      return normalizeHex(color.hex);
    case "system":
      return normalizeHex(color.lastColor ?? "000000");
    case "scheme": {
      const mappedName = context.colorMap[color.scheme] ?? color.scheme;
      if (visited.has(mappedName)) return "#000000";
      const schemeColor = context.theme?.colorScheme?.colors[mappedName];
      return schemeColor !== undefined
        ? resolveColor(context, schemeColor, new Set([...visited, mappedName]))?.hex
        : FALLBACK_SCHEME_COLORS[mappedName];
    }
  }
}

function applyColorTransforms(
  initialHex: string,
  transforms: readonly SourceColorTransform[] | undefined,
): ComputedColor {
  let hex = initialHex;
  let alpha = 1;
  for (const transform of transforms ?? []) {
    switch (transform.kind) {
      case "lumMod": {
        const lumOff = transforms?.find((candidate) => candidate.kind === "lumOff");
        hex = applyLuminance(
          hex,
          Number(transform.value) / 100000,
          Number(lumOff?.value ?? 0) / 100000,
        );
        break;
      }
      case "lumOff":
        if (!transforms?.some((candidate) => candidate.kind === "lumMod")) {
          hex = applyLuminance(hex, 1, Number(transform.value) / 100000);
        }
        break;
      case "tint":
        hex = applyTint(hex, Number(transform.value) / 100000);
        break;
      case "shade":
        hex = applyShade(hex, Number(transform.value) / 100000);
        break;
      case "alpha":
        alpha = Number(transform.value) / 100000;
        break;
    }
  }

  return { hex, alpha };
}

function normalizeHex(hex: string): string {
  const normalized = hex.replace(/^#/, "").toLowerCase();
  return `#${normalized.padStart(6, "0").slice(0, 6)}`;
}

function applyLuminance(hex: string, lumMod: number, lumOff: number): string {
  const { h, s, l } = hexToHsl(hex);
  return hslToHex(h, s, Math.min(1, Math.max(0, l * lumMod + lumOff)));
}

function applyTint(hex: string, tintAmount: number): string {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(
    Math.round(r + (255 - r) * tintAmount),
    Math.round(g + (255 - g) * tintAmount),
    Math.round(b + (255 - b) * tintAmount),
  );
}

function applyShade(hex: string, shadeAmount: number): string {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(
    Math.round(r * shadeAmount),
    Math.round(g * shadeAmount),
    Math.round(b * shadeAmount),
  );
}

function hexToRgb(hex: string): { readonly r: number; readonly g: number; readonly b: number } {
  const normalized = hex.replace("#", "");
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16),
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (value: number) => Math.min(255, Math.max(0, value)).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function hexToHsl(hex: string): { readonly h: number; readonly s: number; readonly l: number } {
  const { r: r255, g: g255, b: b255 } = hexToRgb(hex);
  const r = r255 / 255;
  const g = g255 / 255;
  const b = b255 / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;

  if (max === min) return { h: 0, s: 0, l };

  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h =
    max === r
      ? ((g - b) / d + (g < b ? 6 : 0)) / 6
      : max === g
        ? ((b - r) / d + 2) / 6
        : ((r - g) / d + 4) / 6;

  return { h, s, l };
}

function hslToHex(h: number, s: number, l: number): string {
  if (s === 0) {
    const value = Math.round(l * 255);
    return rgbToHex(value, value, value);
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
  return rgbToHex(
    Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, h) * 255),
    Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
  );
}
