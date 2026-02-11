import type { EffectList, OuterShadow, InnerShadow, Glow, SoftEdge } from "../model/effect.js";
import type { ColorResolver } from "../color/color-resolver.js";

export function parseEffectList(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  effectLstNode: any,
  colorResolver: ColorResolver,
): EffectList | null {
  if (!effectLstNode) return null;

  const outerShadow = parseOuterShadow(effectLstNode.outerShdw, colorResolver);
  const innerShadow = parseInnerShadow(effectLstNode.innerShdw, colorResolver);
  const glow = parseGlow(effectLstNode.glow, colorResolver);
  const softEdge = parseSoftEdge(effectLstNode.softEdge);

  if (!outerShadow && !innerShadow && !glow && !softEdge) {
    return null;
  }

  return { outerShadow, innerShadow, glow, softEdge };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseOuterShadow(node: any, colorResolver: ColorResolver): OuterShadow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    blurRadius: Number(node["@_blurRad"] ?? 0),
    distance: Number(node["@_dist"] ?? 0),
    direction: Number(node["@_dir"] ?? 0) / 60000,
    color,
    alignment: node["@_algn"] ?? "b",
    rotateWithShape: node["@_rotWithShape"] !== "0",
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseInnerShadow(node: any, colorResolver: ColorResolver): InnerShadow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    blurRadius: Number(node["@_blurRad"] ?? 0),
    distance: Number(node["@_dist"] ?? 0),
    direction: Number(node["@_dir"] ?? 0) / 60000,
    color,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseGlow(node: any, colorResolver: ColorResolver): Glow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    radius: Number(node["@_rad"] ?? 0),
    color,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseSoftEdge(node: any): SoftEdge | null {
  if (!node) return null;

  return {
    radius: Number(node["@_rad"] ?? 0),
  };
}
