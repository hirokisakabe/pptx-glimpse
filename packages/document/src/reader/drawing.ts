/**
 * Helpers that read DrawingML colors, fills, lines, and coordinates into
 * PptxSourceModel source types.
 *
 * Theme colors and relationships stay unresolved in source. Transformations such as
 * lumMod and tint are kept unapplied, and the computed view is responsible for resolving
 * them. Unsupported fills such as gradient, pattern, and picture fills are preserved as
 * raw sidecars.
 */

import type {
  RawSidecarId,
  SourceArrowEndpoint,
  SourceArrowSize,
  SourceArrowType,
  SourceBiLevelEffect,
  SourceBlipEffects,
  SourceBlurEffect,
  SourceColor,
  SourceColorChangeEffect,
  SourceColorTransform,
  SourceDashStyle,
  SourceDuotoneEffect,
  SourceEffectList,
  SourceFill,
  SourceGradientStop,
  SourceImageFillTile,
  SourceLineCap,
  SourceLineJoin,
  SourceLumEffect,
  SourceOutline,
  SourceRectangleAlignment,
  SourceTransform,
} from "../source/index.js";
import { asEmu, asOoxmlAngle, asOoxmlPercent, asRelationshipId } from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { parseEnumValue, parseEnumValueWithDefault } from "./ooxml-values.js";
import { makeSidecar } from "./raw-node.js";
import {
  getAttr,
  getChild,
  getChildArray,
  getNamespacedAttr,
  hasChild,
  localName,
  type XmlNode,
} from "./xml.js";

const RECTANGLE_ALIGNMENT_TOKENS: ReadonlySet<SourceRectangleAlignment> = new Set([
  "tl",
  "t",
  "tr",
  "l",
  "ctr",
  "r",
  "bl",
  "b",
  "br",
]);

const COLOR_TRANSFORM_KINDS: ReadonlySet<SourceColorTransform["kind"]> = new Set([
  "lumMod",
  "lumOff",
  "tint",
  "shade",
  "alpha",
]);

/** fill elements that do not interpret typed. Used to exclude raw fill judgment. */
const RAW_FILL_LOCAL_NAMES = ["grpFill"] as const;
const PRESET_COLOR_HEX: Readonly<Record<string, string>> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "008000",
  blue: "0000FF",
  yellow: "FFFF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
};

/**
 * Color elements directly below color preserving elements (`a:solidFill` / `a:bgRef` / `a:rPr`, etc.)
 * Convert (`a:srgbClr` / `a:schemeClr` / `a:sysClr`) to `SourceColor`.
 */
export function parseColorElement(parent: XmlNode | undefined): SourceColor | undefined {
  if (!parent) return undefined;

  const srgb = getChild(parent, "srgbClr");
  if (srgb) {
    const hex = getAttr(srgb, "val");
    if (hex !== undefined) {
      return withTransforms({ kind: "srgb", hex: hex.toUpperCase() }, srgb);
    }
  }

  const scheme = getChild(parent, "schemeClr");
  if (scheme) {
    const name = getAttr(scheme, "val");
    if (name !== undefined) {
      return withTransforms({ kind: "scheme", scheme: name }, scheme);
    }
  }

  const sys = getChild(parent, "sysClr");
  if (sys) {
    const val = getAttr(sys, "val");
    if (val !== undefined) {
      const lastClr = getAttr(sys, "lastClr");
      return withTransforms(
        {
          kind: "system",
          value: val,
          ...(lastClr !== undefined ? { lastColor: lastClr.toUpperCase() } : {}),
        },
        sys,
      );
    }
  }

  return undefined;
}

export function parseEffectList(effectList: XmlNode | undefined): SourceEffectList | undefined {
  if (effectList === undefined) return undefined;
  const outerShadow = parseOuterShadow(getChild(effectList, "outerShdw"));
  const innerShadow = parseInnerShadow(getChild(effectList, "innerShdw"));
  const glow = parseGlow(getChild(effectList, "glow"));
  const softEdge = parseSoftEdge(getChild(effectList, "softEdge"));
  const parsed: SourceEffectList = {
    ...(outerShadow !== undefined ? { outerShadow } : {}),
    ...(innerShadow !== undefined ? { innerShadow } : {}),
    ...(glow !== undefined ? { glow } : {}),
    ...(softEdge !== undefined ? { softEdge } : {}),
  };
  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

export function parseBlipEffects(blip: XmlNode | undefined): SourceBlipEffects | undefined {
  if (blip === undefined) return undefined;
  const grayscale = getChild(blip, "grayscl") !== undefined;
  const biLevel = parseBiLevel(getChild(blip, "biLevel"));
  const blur = parseBlurEffect(getChild(blip, "blur"));
  const lum = parseLumEffect(getChild(blip, "lum"));
  const duotone = parseDuotoneEffect(getChild(blip, "duotone"));
  const clrChange = parseColorChangeEffect(getChild(blip, "clrChange"));
  const parsed: SourceBlipEffects = {
    grayscale,
    ...(biLevel !== undefined ? { biLevel } : {}),
    ...(blur !== undefined ? { blur } : {}),
    ...(lum !== undefined ? { lum } : {}),
    ...(duotone !== undefined ? { duotone } : {}),
    ...(clrChange !== undefined ? { clrChange } : {}),
  };
  return grayscale ||
    biLevel !== undefined ||
    blur !== undefined ||
    lum !== undefined ||
    duotone !== undefined ||
    clrChange !== undefined
    ? parsed
    : undefined;
}

function parseOuterShadow(node: XmlNode | undefined): SourceEffectList["outerShadow"] | undefined {
  if (node === undefined) return undefined;
  const color = parseColorElement(node);
  if (color === undefined) return undefined;
  return {
    blurRadius: asEmu(numericAttr(node, "blurRad") ?? 0),
    distance: asEmu(numericAttr(node, "dist") ?? 0),
    direction: asOoxmlAngle(numericAttr(node, "dir") ?? 0),
    color,
    alignment: parseRectangleAlignment(getAttr(node, "algn"), "b"),
    rotateWithShape: getAttr(node, "rotWithShape") !== "0",
  };
}

function parseInnerShadow(node: XmlNode | undefined): SourceEffectList["innerShadow"] | undefined {
  if (node === undefined) return undefined;
  const color = parseColorElement(node);
  if (color === undefined) return undefined;
  return {
    blurRadius: asEmu(numericAttr(node, "blurRad") ?? 0),
    distance: asEmu(numericAttr(node, "dist") ?? 0),
    direction: asOoxmlAngle(numericAttr(node, "dir") ?? 0),
    color,
  };
}

function parseGlow(node: XmlNode | undefined): SourceEffectList["glow"] | undefined {
  if (node === undefined) return undefined;
  const color = parseColorElement(node);
  if (color === undefined) return undefined;
  return {
    radius: asEmu(numericAttr(node, "rad") ?? 0),
    color,
  };
}

function parseSoftEdge(node: XmlNode | undefined): SourceEffectList["softEdge"] | undefined {
  if (node === undefined) return undefined;
  return {
    radius: asEmu(numericAttr(node, "rad") ?? 0),
  };
}

function parseBiLevel(node: XmlNode | undefined): SourceBiLevelEffect | undefined {
  if (node === undefined) return undefined;
  return {
    threshold: (numericAttr(node, "thresh") ?? 50000) / 100000,
  };
}

function parseBlurEffect(node: XmlNode | undefined): SourceBlurEffect | undefined {
  if (node === undefined) return undefined;
  return {
    radius: asEmu(numericAttr(node, "rad") ?? 0),
    grow: getAttr(node, "grow") !== "0",
  };
}

function parseLumEffect(node: XmlNode | undefined): SourceLumEffect | undefined {
  if (node === undefined) return undefined;
  return {
    brightness: (numericAttr(node, "bright") ?? 0) / 100000,
    contrast: (numericAttr(node, "contrast") ?? 0) / 100000,
  };
}

function parseDuotoneEffect(node: XmlNode | undefined): SourceDuotoneEffect | undefined {
  if (node === undefined) return undefined;
  const colors = collectColorChildren(node);
  if (colors.length < 2) return undefined;
  return { color1: colors[0], color2: colors[1] };
}

function parseColorChangeEffect(node: XmlNode | undefined): SourceColorChangeEffect | undefined {
  if (node === undefined) return undefined;
  const from = firstColorChild(getChild(node, "clrFrom"));
  const to = firstColorChild(getChild(node, "clrTo"));
  return from !== undefined && to !== undefined ? { from, to } : undefined;
}

function collectColorChildren(parent: XmlNode): SourceColor[] {
  const colors: SourceColor[] = [];
  for (const [key, value] of Object.entries(parent)) {
    if (key.startsWith("@_")) continue;
    const nodes = Array.isArray(value) ? value : [value];
    for (const node of nodes) {
      if (!isXmlNode(node)) continue;
      const color = parseColorChild(key, node);
      if (color !== undefined) colors.push(color);
    }
  }
  return colors;
}

function firstColorChild(parent: XmlNode | undefined): SourceColor | undefined {
  return parent !== undefined ? collectColorChildren(parent)[0] : undefined;
}

function parseColorChild(key: string, node: XmlNode): SourceColor | undefined {
  const name = localName(key);
  return name === "prstClr" ? parsePresetColor(node) : parseColorElement({ [name]: node });
}

function isXmlNode(value: unknown): value is XmlNode {
  return typeof value === "object" && value !== null;
}

function parsePresetColor(node: XmlNode): SourceColor | undefined {
  const value = getAttr(node, "val");
  const hex = value !== undefined ? PRESET_COLOR_HEX[value] : undefined;
  return hex !== undefined ? withTransforms({ kind: "srgb", hex }, node) : undefined;
}

/**
 * Read the fill immediately below `a:spPr` / `a:ln` / `p:bgPr` etc. `a:solidFill` /
 * Convert `a:noFill` to typed and gradient / pattern / picture fill to raw.
 */
export function parseFill(
  parent: XmlNode | undefined,
  nextId: () => RawSidecarId,
): SourceFill | undefined {
  if (!parent) return undefined;

  const solid = getChild(parent, "solidFill");
  if (solid) {
    const color = parseColorElement(solid);
    if (color) return { kind: "solid", color };
    return { kind: "raw", raw: makeSidecar("a:solidFill", solid, nextId) };
  }

  if (hasChild(parent, "noFill")) return { kind: "none" };

  const grad = getChild(parent, "gradFill");
  if (grad) {
    const fill = parseGradientFill(grad);
    return fill ?? { kind: "raw", raw: makeSidecar("a:gradFill", grad, nextId) };
  }

  const blip = getChild(parent, "blipFill");
  if (blip) {
    return parseBlipFill(blip) ?? { kind: "raw", raw: makeSidecar("a:blipFill", blip, nextId) };
  }

  const pattern = getChild(parent, "pattFill");
  if (pattern) {
    const fill = parsePatternFill(pattern);
    return fill ?? { kind: "raw", raw: makeSidecar("a:pattFill", pattern, nextId) };
  }

  for (const name of RAW_FILL_LOCAL_NAMES) {
    if (hasChild(parent, name)) {
      return { kind: "raw", raw: makeSidecar(`a:${name}`, getChild(parent, name) ?? {}, nextId) };
    }
  }

  return undefined;
}

function parseGradientFill(grad: XmlNode): SourceFill | undefined {
  const stops: SourceGradientStop[] = [];
  for (const gs of getChildArray(getChild(grad, "gsLst"), "gs")) {
    const color = parseColorElement(gs);
    if (color === undefined) continue;
    stops.push({
      position: (numericAttr(gs, "pos") ?? 0) / 100000,
      color,
    });
  }
  if (stops.length === 0) return undefined;

  const path = getChild(grad, "path");
  if (path !== undefined) {
    const fillToRect = getChild(path, "fillToRect");
    const l = numericAttr(fillToRect, "l") ?? 0;
    const t = numericAttr(fillToRect, "t") ?? 0;
    const r = numericAttr(fillToRect, "r") ?? 0;
    const b = numericAttr(fillToRect, "b") ?? 0;
    return {
      kind: "gradient",
      gradientType: "radial",
      stops,
      centerX: (l + (100000 - r)) / 2 / 100000,
      centerY: (t + (100000 - b)) / 2 / 100000,
    };
  }

  return {
    kind: "gradient",
    gradientType: "linear",
    stops,
    angle: asOoxmlAngle(numericAttr(getChild(grad, "lin"), "ang") ?? 0),
  };
}

function parseBlipFill(blipFill: XmlNode): SourceFill | undefined {
  const blip = getChild(blipFill, "blip");
  const embed = getNamespacedAttr(blip, "embed");
  if (embed === undefined) return undefined;
  const tile = parseImageFillTile(getChild(blipFill, "tile"));
  return {
    kind: "image",
    blipRelationshipId: asRelationshipId(embed),
    ...(tile !== undefined ? { tile } : {}),
  };
}

export function parseImageFillTile(tile: XmlNode | undefined): SourceImageFillTile | undefined {
  if (tile === undefined) return undefined;
  const flip = getAttr(tile, "flip") ?? "none";
  return {
    tx: asEmu(numericAttr(tile, "tx") ?? 0),
    ty: asEmu(numericAttr(tile, "ty") ?? 0),
    sx: (numericAttr(tile, "sx") ?? 100000) / 100000,
    sy: (numericAttr(tile, "sy") ?? 100000) / 100000,
    flip: flip === "x" || flip === "y" || flip === "xy" ? flip : "none",
    align: parseRectangleAlignment(getAttr(tile, "algn"), "tl"),
  };
}

export function parseRectangleAlignment(
  value: string | undefined,
  fallback: SourceRectangleAlignment,
): SourceRectangleAlignment {
  return parseEnumValueWithDefault(value, RECTANGLE_ALIGNMENT_TOKENS, fallback);
}

function parsePatternFill(pattern: XmlNode): SourceFill | undefined {
  const foregroundColor = parseColorElement(getChild(pattern, "fgClr"));
  const backgroundColor = parseColorElement(getChild(pattern, "bgClr"));
  if (foregroundColor === undefined || backgroundColor === undefined) return undefined;
  return {
    kind: "pattern",
    preset: getAttr(pattern, "prst") ?? "ltDnDiag",
    foregroundColor,
    backgroundColor,
  };
}

/** Read `a:ln`. Minimal representation of width (EMU) and solid line color only. */
export function parseOutline(
  spPr: XmlNode | undefined,
  nextId: () => RawSidecarId,
): SourceOutline | undefined {
  const ln = getChild(spPr, "ln");
  return parseLine(ln, nextId);
}

/** Read line nodes such as `a:ln` / `a:lnL` / `a:lnR`. */
export function parseLine(
  ln: XmlNode | undefined,
  nextId: () => RawSidecarId,
): SourceOutline | undefined {
  if (!ln) return undefined;

  const width = numericAttr(ln, "w");
  const fill = parseFill(ln, nextId);
  const dashStyle = parseDashStyle(getChild(ln, "prstDash"));
  const customDash = parseCustomDash(ln);
  const lineCap = parseLineCap(getAttr(ln, "cap"));
  const lineJoin = parseLineJoin(ln);
  const headEnd = parseArrowEndpoint(getChild(ln, "headEnd"));
  const tailEnd = parseArrowEndpoint(getChild(ln, "tailEnd"));
  return {
    ...(width !== undefined ? { width: asEmu(width) } : {}),
    ...(fill !== undefined ? { fill } : {}),
    ...(dashStyle !== undefined ? { dashStyle } : {}),
    ...(customDash !== undefined ? { customDash } : {}),
    ...(lineCap !== undefined ? { lineCap } : {}),
    ...(lineJoin !== undefined ? { lineJoin } : {}),
    ...(headEnd !== undefined ? { headEnd } : {}),
    ...(tailEnd !== undefined ? { tailEnd } : {}),
  };
}

/** Read `a:xfrm` and set offset / extent / rotation / flip to `SourceTransform`. */
export function parseTransform(spPr: XmlNode | undefined): SourceTransform | undefined {
  const xfrm = getChild(spPr, "xfrm");
  if (!xfrm) return undefined;

  const off = getChild(xfrm, "off");
  const ext = getChild(xfrm, "ext");
  const offsetX = numericAttr(off, "x");
  const offsetY = numericAttr(off, "y");
  const width = numericAttr(ext, "cx");
  const height = numericAttr(ext, "cy");
  if (
    offsetX === undefined ||
    offsetY === undefined ||
    width === undefined ||
    height === undefined
  ) {
    return undefined;
  }

  const rotation = numericAttr(xfrm, "rot");
  const flipH = getAttr(xfrm, "flipH");
  const flipV = getAttr(xfrm, "flipV");
  return {
    offsetX: asEmu(offsetX),
    offsetY: asEmu(offsetY),
    width: asEmu(width),
    height: asEmu(height),
    ...(rotation !== undefined ? { rotation: asOoxmlAngle(rotation) } : {}),
    ...(isTrue(flipH) ? { flipHorizontal: true } : {}),
    ...(isTrue(flipV) ? { flipVertical: true } : {}),
  };
}

function withTransforms(base: SourceColor, colorNode: XmlNode): SourceColor {
  const transforms: SourceColorTransform[] = [];
  for (const key of Object.keys(colorNode)) {
    if (key.startsWith("@_")) continue;
    const kind = localName(key);
    const transformKind = parseEnumValue(kind, COLOR_TRANSFORM_KINDS);
    if (transformKind === undefined) continue;
    const value = colorNode[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      const raw = getAttr(unsafeOoxmlBoundaryAssertion<XmlNode>(item), "val");
      if (raw === undefined) continue;
      const numeric = Number(raw);
      if (!Number.isFinite(numeric)) continue;
      transforms.push({
        kind: transformKind,
        value: asOoxmlPercent(numeric),
      });
    }
  }
  return transforms.length > 0 ? { ...base, transforms } : base;
}

/** Extract numeric attributes. Missing/non-numeric values are undefined. */
export function numericAttr(node: XmlNode | undefined, name: string): number | undefined {
  const raw = getAttr(node, name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

/** Determine OOXML boolean attribute (`1` / `0` / `true` / `false`). */
export function isTrue(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function parseDashStyle(prstDash: XmlNode | undefined): SourceDashStyle | undefined {
  const value = getAttr(prstDash, "val");
  return parseEnumValue(value, DASH_STYLES);
}

function parseCustomDash(ln: XmlNode): number[] | undefined {
  const custDash = getChild(ln, "custDash");
  const segments = getChildArray(custDash, "ds");
  if (segments.length === 0) return undefined;
  return segments.flatMap((segment) => [
    (numericAttr(segment, "d") ?? 100000) / 100000,
    (numericAttr(segment, "sp") ?? 100000) / 100000,
  ]);
}

const LINE_CAP_MAP: Record<string, SourceLineCap> = {
  flat: "butt",
  sq: "square",
  rnd: "round",
};

function parseLineCap(value: string | undefined): SourceLineCap | undefined {
  return value !== undefined ? LINE_CAP_MAP[value] : undefined;
}

function parseLineJoin(ln: XmlNode): SourceLineJoin | undefined {
  if (hasChild(ln, "round")) return "round";
  if (hasChild(ln, "bevel")) return "bevel";
  if (hasChild(ln, "miter")) return "miter";
  return undefined;
}

function parseArrowEndpoint(node: XmlNode | undefined): SourceArrowEndpoint | undefined {
  if (!node) return undefined;
  const type = parseEnumValue(getAttr(node, "type"), ARROW_TYPES);
  if (type === undefined) return undefined;
  const width = getAttr(node, "w") ?? "med";
  const length = getAttr(node, "len") ?? "med";
  return {
    type,
    width: parseEnumValueWithDefault(width, ARROW_SIZES, "med"),
    length: parseEnumValueWithDefault(length, ARROW_SIZES, "med"),
  };
}

const DASH_STYLES: ReadonlySet<SourceDashStyle> = new Set([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "sysDash",
  "sysDot",
]);

const ARROW_TYPES: ReadonlySet<SourceArrowType> = new Set([
  "triangle",
  "stealth",
  "diamond",
  "oval",
  "arrow",
]);

const ARROW_SIZES: ReadonlySet<SourceArrowSize> = new Set(["sm", "med", "lg"]);
