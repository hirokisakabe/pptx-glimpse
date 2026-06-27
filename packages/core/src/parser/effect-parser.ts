import type { EffectList, Glow, InnerShadow, OuterShadow, SoftEdge } from "@pptx-glimpse/renderer";
import { asEmu } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import type { XmlNode } from "./xml-parser.js";

export function parseEffectList(
  effectLstNode: XmlNode,
  colorResolver: ColorResolver,
): EffectList | null {
  if (!effectLstNode) return null;

  const outerShadow = parseOuterShadow(
    unsafeTypeAssertion<XmlNode>(effectLstNode.outerShdw),
    colorResolver,
  );
  const innerShadow = parseInnerShadow(
    unsafeTypeAssertion<XmlNode>(effectLstNode.innerShdw),
    colorResolver,
  );
  const glow = parseGlow(unsafeTypeAssertion<XmlNode>(effectLstNode.glow), colorResolver);
  const softEdge = parseSoftEdge(unsafeTypeAssertion<XmlNode>(effectLstNode.softEdge));

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
    blurRadius: asEmu(Number(node["@_blurRad"] ?? 0)),
    distance: asEmu(Number(node["@_dist"] ?? 0)),
    direction: Number(node["@_dir"] ?? 0) / 60000,
    color,
    alignment: unsafeTypeAssertion<string | undefined>(node["@_algn"]) ?? "b",
    rotateWithShape: node["@_rotWithShape"] !== "0",
  };
}

function parseInnerShadow(node: XmlNode, colorResolver: ColorResolver): InnerShadow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    blurRadius: asEmu(Number(node["@_blurRad"] ?? 0)),
    distance: asEmu(Number(node["@_dist"] ?? 0)),
    direction: Number(node["@_dir"] ?? 0) / 60000,
    color,
  };
}

function parseGlow(node: XmlNode, colorResolver: ColorResolver): Glow | null {
  if (!node) return null;

  const color = colorResolver.resolve(node);
  if (!color) return null;

  return {
    radius: asEmu(Number(node["@_rad"] ?? 0)),
    color,
  };
}

function parseSoftEdge(node: XmlNode): SoftEdge | null {
  if (!node) return null;

  return {
    radius: asEmu(Number(node["@_rad"] ?? 0)),
  };
}
