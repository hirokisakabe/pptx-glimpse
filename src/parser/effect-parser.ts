import type { EffectList, OuterShadow, InnerShadow, Glow, SoftEdge } from "../model/effect.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { XmlNode } from "./xml-parser.js";

export function parseEffectList(
  effectLstNode: XmlNode,
  colorResolver: ColorResolver,
): EffectList | null {
  if (!effectLstNode) return null;

  const outerShadow = parseOuterShadow(effectLstNode.outerShdw as XmlNode, colorResolver);
  const innerShadow = parseInnerShadow(effectLstNode.innerShdw as XmlNode, colorResolver);
  const glow = parseGlow(effectLstNode.glow as XmlNode, colorResolver);
  const softEdge = parseSoftEdge(effectLstNode.softEdge as XmlNode);

  if (!outerShadow && !innerShadow && !glow && !softEdge) {
    return null;
  }

  return { outerShadow, innerShadow, glow, softEdge };
}

function parseOuterShadow(node: XmlNode, colorResolver: ColorResolver): OuterShadow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    blurRadius: Number(node["@_blurRad"] ?? 0),
    distance: Number(node["@_dist"] ?? 0),
    direction: Number(node["@_dir"] ?? 0) / 60000,
    color,
    alignment: (node["@_algn"] as string | undefined) ?? "b",
    rotateWithShape: node["@_rotWithShape"] !== "0",
  };
}

function parseInnerShadow(node: XmlNode, colorResolver: ColorResolver): InnerShadow | null {
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

function parseGlow(node: XmlNode, colorResolver: ColorResolver): Glow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    radius: Number(node["@_rad"] ?? 0),
    color,
  };
}

function parseSoftEdge(node: XmlNode): SoftEdge | null {
  if (!node) return null;

  return {
    radius: Number(node["@_rad"] ?? 0),
  };
}
