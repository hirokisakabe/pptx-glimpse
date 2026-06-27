import type {
  BiLevelEffect,
  BlipEffects,
  BlurEffect,
  ClrChangeEffect,
  DuotoneEffect,
  LumEffect,
} from "@pptx-glimpse/renderer";
import type { ResolvedColor } from "@pptx-glimpse/renderer";
import { asEmu } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import type { XmlNode } from "./xml-parser.js";

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
  const biLevel = parseBiLevel(unsafeTypeAssertion<XmlNode | undefined>(blipNode.biLevel));
  const blur = parseBlur(unsafeTypeAssertion<XmlNode | undefined>(blipNode.blur));
  const lum = parseLum(unsafeTypeAssertion<XmlNode | undefined>(blipNode.lum));
  const duotone = parseDuotone(
    unsafeTypeAssertion<XmlNode | undefined>(blipNode.duotone),
    colorResolver,
  );
  const clrChange = parseClrChange(
    unsafeTypeAssertion<XmlNode | undefined>(blipNode.clrChange),
    colorResolver,
  );

  if (!grayscale && !biLevel && !blur && !lum && !duotone && !clrChange) {
    return null;
  }

  return { grayscale, biLevel, blur, lum, duotone, clrChange };
}

function parseBiLevel(node: XmlNode | undefined): BiLevelEffect | null {
  if (!node) return null;
  const threshold = Number(node["@_thresh"] ?? 50000) / 100000;
  return { threshold };
}

function parseBlur(node: XmlNode | undefined): BlurEffect | null {
  if (!node) return null;
  return {
    radius: asEmu(Number(node["@_rad"] ?? 0)),
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
      const resolved = resolveColorNode(key, unsafeTypeAssertion<XmlNode>(cn), colorResolver);
      if (resolved) colors.push(resolved);
    }
  }

  if (colors.length < 2) return null;
  return { color1: colors[0], color2: colors[1] };
}

function parseClrChange(
  node: XmlNode | undefined,
  colorResolver: ColorResolver,
): ClrChangeEffect | null {
  if (!node) return null;

  const clrFrom = unsafeTypeAssertion<XmlNode | undefined>(node.clrFrom);
  const clrTo = unsafeTypeAssertion<XmlNode | undefined>(node.clrTo);
  if (!clrFrom || !clrTo) return null;

  const from = colorResolver.resolve(clrFrom);
  const to = colorResolver.resolve(clrTo);
  if (!from || !to) return null;

  return { clrFrom: from, clrTo: to };
}

function resolveColorNode(
  key: string,
  node: XmlNode,
  colorResolver: ColorResolver,
): ResolvedColor | null {
  if (key === "prstClr") {
    const val = unsafeTypeAssertion<string | undefined>(node["@_val"]);
    const hex = val ? PRESET_COLORS[val] : undefined;
    if (!hex) return null;
    return { hex, alpha: 1 };
  }
  return colorResolver.resolve({ [key]: node });
}
