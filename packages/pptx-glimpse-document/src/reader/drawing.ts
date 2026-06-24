/**
 * DrawingML の色・塗り・線・座標変換を CleanDoc source 型へ読み取るヘルパー。
 *
 * source では theme color / relationship を未解決のまま保持する
 * (`docs/cleandoc-source-computed-view.md`)。lumMod / tint 等の変換は適用せず
 * そのまま保存し、解決は computed view の責務とする。未対応の塗り (gradient /
 * pattern / picture fill) は raw sidecar として保存する。
 */

import type {
  RawSidecarId,
  SourceArrowEndpoint,
  SourceArrowSize,
  SourceArrowType,
  SourceColor,
  SourceColorTransform,
  SourceDashStyle,
  SourceFill,
  SourceLineCap,
  SourceLineJoin,
  SourceOutline,
  SourceTransform,
} from "../source/index.js";
import { asEmu, asOoxmlAngle, asOoxmlPercent } from "../source/index.js";
import { makeSidecar } from "./raw-node.js";
import { getAttr, getChild, getChildArray, hasChild, localName, type XmlNode } from "./xml.js";

const COLOR_TRANSFORM_KINDS: ReadonlySet<string> = new Set([
  "lumMod",
  "lumOff",
  "tint",
  "shade",
  "alpha",
]);

/** typed に解釈する fill 要素。raw fill 判定で除外するために使う。 */
const RAW_FILL_LOCAL_NAMES = ["gradFill", "blipFill", "pattFill", "grpFill"] as const;

/**
 * 色保持要素 (`a:solidFill` / `a:bgRef` / `a:rPr` 等) の直下にある色要素
 * (`a:srgbClr` / `a:schemeClr` / `a:sysClr`) を `SourceColor` に変換する。
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

/**
 * `a:spPr` / `a:ln` / `p:bgPr` 等の直下にある塗りを読む。`a:solidFill` /
 * `a:noFill` を typed に、gradient / pattern / picture fill を raw に変換する。
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

  for (const name of RAW_FILL_LOCAL_NAMES) {
    const node = getChild(parent, name);
    if (node) return { kind: "raw", raw: makeSidecar(`a:${name}`, node, nextId) };
  }

  return undefined;
}

/** `a:ln` を読む。幅 (EMU) と solid line color のみの最小表現。 */
export function parseOutline(
  spPr: XmlNode | undefined,
  nextId: () => RawSidecarId,
): SourceOutline | undefined {
  const ln = getChild(spPr, "ln");
  return parseLine(ln, nextId);
}

/** `a:ln` / `a:lnL` / `a:lnR` 等の line node を読む。 */
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

/** `a:xfrm` を読み、offset / extent / rotation / flip を `SourceTransform` に。 */
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
    if (!COLOR_TRANSFORM_KINDS.has(kind)) continue;
    const value = colorNode[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      const raw = getAttr(item as XmlNode, "val");
      if (raw === undefined) continue;
      const numeric = Number(raw);
      if (!Number.isFinite(numeric)) continue;
      transforms.push({
        kind: kind as SourceColorTransform["kind"],
        value: asOoxmlPercent(numeric),
      });
    }
  }
  return transforms.length > 0 ? { ...base, transforms } : base;
}

/** 数値属性を取り出す。欠落・非数値は undefined。 */
export function numericAttr(node: XmlNode | undefined, name: string): number | undefined {
  const raw = getAttr(node, name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

/** OOXML boolean 属性 (`1` / `0` / `true` / `false`) を判定する。 */
export function isTrue(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function parseDashStyle(prstDash: XmlNode | undefined): SourceDashStyle | undefined {
  const value = getAttr(prstDash, "val");
  if (value === undefined) return undefined;
  return DASH_STYLES.has(value) ? (value as SourceDashStyle) : undefined;
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
  const type = (getAttr(node, "type") ?? "none") as SourceArrowType | "none";
  if (type === "none") return undefined;
  if (!ARROW_TYPES.has(type)) return undefined;
  const width = getAttr(node, "w") ?? "med";
  const length = getAttr(node, "len") ?? "med";
  return {
    type,
    width: ARROW_SIZES.has(width) ? (width as SourceArrowSize) : "med",
    length: ARROW_SIZES.has(length) ? (length as SourceArrowSize) : "med",
  };
}

const DASH_STYLES: ReadonlySet<string> = new Set([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "sysDash",
  "sysDot",
]);

const ARROW_TYPES: ReadonlySet<string> = new Set([
  "triangle",
  "stealth",
  "diamond",
  "oval",
  "arrow",
]);

const ARROW_SIZES: ReadonlySet<string> = new Set(["sm", "med", "lg"]);
