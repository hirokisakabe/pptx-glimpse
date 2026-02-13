import type {
  BlipEffects,
  BiLevelEffect,
  BlurEffect,
  LumEffect,
  DuotoneEffect,
} from "../model/effect.js";
import type { ResolvedColor } from "../model/fill.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { XmlNode } from "./xml-parser.js";
import { warn } from "../warning-logger.js";

const PRESET_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
};

export function parseBlipEffects(
  blipNode: XmlNode,
  colorResolver: ColorResolver,
): BlipEffects | null {
  if (!blipNode) return null;

  const grayscale = blipNode.grayscl !== undefined;
  const biLevel = parseBiLevel(blipNode.biLevel as XmlNode | undefined);
  const blur = parseBlur(blipNode.blur as XmlNode | undefined);
  const lum = parseLum(blipNode.lum as XmlNode | undefined);
  const duotone = parseDuotone(blipNode.duotone as XmlNode | undefined, colorResolver);

  if (blipNode.clrChange !== undefined) {
    warn("blip.clrChange", "color change effect not implemented");
  }

  if (!grayscale && !biLevel && !blur && !lum && !duotone) {
    return null;
  }

  return { grayscale, biLevel, blur, lum, duotone };
}

function parseBiLevel(node: XmlNode | undefined): BiLevelEffect | null {
  if (!node) return null;
  const threshold = Number(node["@_thresh"] ?? 50000) / 100000;
  return { threshold };
}

function parseBlur(node: XmlNode | undefined): BlurEffect | null {
  if (!node) return null;
  return {
    radius: Number(node["@_rad"] ?? 0),
    grow: node["@_grow"] !== "0",
  };
}

function parseLum(node: XmlNode | undefined): LumEffect | null {
  if (!node) return null;
  return {
    brightness: Number(node["@_bright"] ?? 0) / 100000,
    contrast: Number(node["@_contrast"] ?? 0) / 100000,
  };
}

function parseDuotone(
  node: XmlNode | undefined,
  colorResolver: ColorResolver,
): DuotoneEffect | null {
  if (!node) return null;

  const colors: ResolvedColor[] = [];

  for (const key of ["prstClr", "srgbClr", "schemeClr", "sysClr"]) {
    const colorNodes = node[key];
    if (!colorNodes) continue;
    const nodes = Array.isArray(colorNodes) ? colorNodes : [colorNodes];
    for (const cn of nodes) {
      const resolved = resolveColorNode(key, cn as XmlNode, colorResolver);
      if (resolved) colors.push(resolved);
    }
  }

  if (colors.length < 2) return null;
  return { color1: colors[0], color2: colors[1] };
}

function resolveColorNode(
  key: string,
  node: XmlNode,
  colorResolver: ColorResolver,
): ResolvedColor | null {
  if (key === "prstClr") {
    const val = node["@_val"] as string | undefined;
    const hex = val ? PRESET_COLORS[val] : undefined;
    if (!hex) return null;
    return { hex, alpha: 1 };
  }
  return colorResolver.resolve({ [key]: node });
}
