/**
 * slide / slideLayout / slideMaster / theme part を CleanDoc source の typed
 * node へ読み取る。
 *
 * 各 part は分離したまま保持し、cascade (master → layout → slide) の解決は
 * computed view の責務とする (`docs/cleandoc-source-computed-view.md`)。本
 * モジュールは解決に必要な素材 (背景 / clrMap / clrMapOvr / theme scheme /
 * showMasterSp) と shape tree を source-local に読み出す。
 */

import type {
  PartPath,
  RawSidecarId,
  SourceBackground,
  SourceColorMap,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
  SourceThemeColorScheme,
  SourceThemeFontScheme,
} from "../source/index.js";
import { type SourceColor } from "../source/index.js";
import { isTrue, numericAttr, parseColorElement, parseFill } from "./drawing.js";
import { makeSidecar } from "./raw-node.js";
import { parseShapeTree } from "./shape-tree.js";
import { getAttr, getAttrs, getChild, type XmlNode, type XmlOrderedNode } from "./xml.js";

const THEME_COLOR_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

/** parse 済み slide root (`p:sld`) から `SourceSlide` を組み立てる。 */
export function parseSlide(
  root: XmlNode | undefined,
  partPath: PartPath,
  layoutPartPath: PartPath,
  nextId: () => RawSidecarId,
  orderedSpTree?: readonly XmlOrderedNode[],
): SourceSlide {
  const cSld = getChild(root, "cSld");
  const background = parseBackground(getChild(cSld, "bg"), nextId);
  const colorMapOverride = parseColorMapOverride(getChild(root, "clrMapOvr"));
  const showMasterShapes = booleanAttr(root, "showMasterSp");

  return {
    partPath,
    layoutPartPath,
    ...(background !== undefined ? { background } : {}),
    ...(colorMapOverride !== undefined ? { colorMapOverride } : {}),
    ...(showMasterShapes !== undefined ? { showMasterShapes } : {}),
    shapes: parseShapeTree(getChild(cSld, "spTree"), partPath, nextId, orderedSpTree),
    handle: { partPath },
  };
}

/** parse 済み layout root (`p:sldLayout`) から `SourceSlideLayout` を組み立てる。 */
export function parseSlideLayout(
  root: XmlNode | undefined,
  partPath: PartPath,
  masterPartPath: PartPath,
  nextId: () => RawSidecarId,
  orderedSpTree?: readonly XmlOrderedNode[],
): SourceSlideLayout {
  const cSld = getChild(root, "cSld");
  const type = getAttr(root, "type");
  const background = parseBackground(getChild(cSld, "bg"), nextId);
  const colorMapOverride = parseColorMapOverride(getChild(root, "clrMapOvr"));
  const showMasterShapes = booleanAttr(root, "showMasterSp");

  return {
    partPath,
    masterPartPath,
    ...(type !== undefined ? { type } : {}),
    ...(background !== undefined ? { background } : {}),
    ...(colorMapOverride !== undefined ? { colorMapOverride } : {}),
    ...(showMasterShapes !== undefined ? { showMasterShapes } : {}),
    shapes: parseShapeTree(getChild(cSld, "spTree"), partPath, nextId, orderedSpTree),
    handle: { partPath },
  };
}

/** parse 済み master root (`p:sldMaster`) から `SourceSlideMaster` を組み立てる。 */
export function parseSlideMaster(
  root: XmlNode | undefined,
  partPath: PartPath,
  themePartPath: PartPath | undefined,
  layoutPartPaths: readonly PartPath[],
  nextId: () => RawSidecarId,
  orderedSpTree?: readonly XmlOrderedNode[],
): SourceSlideMaster {
  const cSld = getChild(root, "cSld");
  const background = parseBackground(getChild(cSld, "bg"), nextId);
  const colorMap = parseColorMap(getChild(root, "clrMap"));

  return {
    partPath,
    ...(themePartPath !== undefined ? { themePartPath } : {}),
    layoutPartPaths,
    ...(background !== undefined ? { background } : {}),
    ...(colorMap !== undefined ? { colorMap } : {}),
    shapes: parseShapeTree(getChild(cSld, "spTree"), partPath, nextId, orderedSpTree),
    handle: { partPath },
  };
}

/** parse 済み theme root (`a:theme`) から `SourceTheme` を組み立てる。 */
export function parseTheme(root: XmlNode | undefined, partPath: PartPath): SourceTheme {
  const name = getAttr(root, "name");
  const themeElements = getChild(root, "themeElements");
  const colorScheme = parseColorScheme(getChild(themeElements, "clrScheme"));
  const fontScheme = parseFontScheme(getChild(themeElements, "fontScheme"));

  return {
    partPath,
    ...(name !== undefined ? { name } : {}),
    ...(colorScheme !== undefined ? { colorScheme } : {}),
    ...(fontScheme !== undefined ? { fontScheme } : {}),
    handle: { partPath },
  };
}

function parseBackground(
  bg: XmlNode | undefined,
  nextId: () => RawSidecarId,
): SourceBackground | undefined {
  if (!bg) return undefined;

  const bgPr = getChild(bg, "bgPr");
  if (bgPr) {
    const fill = parseFill(bgPr, nextId);
    if (fill !== undefined) return { kind: "fill", fill };
    return { kind: "raw", raw: makeSidecar("p:bg", bg, nextId) };
  }

  const bgRef = getChild(bg, "bgRef");
  if (bgRef) {
    const color = parseColorElement(bgRef);
    const index = numericAttr(bgRef, "idx");
    if (color !== undefined) {
      return { kind: "styleReference", index: index ?? 0, color };
    }
  }

  return { kind: "raw", raw: makeSidecar("p:bg", bg, nextId) };
}

function parseColorMap(clrMap: XmlNode | undefined): SourceColorMap | undefined {
  if (!clrMap) return undefined;
  const mapping = getAttrs(clrMap);
  return Object.keys(mapping).length > 0 ? { mapping } : undefined;
}

/**
 * `p:clrMapOvr` を読む。`a:overrideClrMapping` があればその mapping を返し、
 * `a:masterClrMapping` (= master の clrMap を踏襲) の場合は override 無しとして
 * undefined を返す。
 */
function parseColorMapOverride(clrMapOvr: XmlNode | undefined): SourceColorMap | undefined {
  if (!clrMapOvr) return undefined;
  const override = getChild(clrMapOvr, "overrideClrMapping");
  if (!override) return undefined;
  const mapping = getAttrs(override);
  return Object.keys(mapping).length > 0 ? { mapping } : undefined;
}

function parseColorScheme(clrScheme: XmlNode | undefined): SourceThemeColorScheme | undefined {
  if (!clrScheme) return undefined;
  const colors: Record<string, SourceColor> = {};
  for (const slot of THEME_COLOR_SLOTS) {
    const color = parseColorElement(getChild(clrScheme, slot));
    if (color !== undefined) colors[slot] = color;
  }
  return Object.keys(colors).length > 0 ? { colors } : undefined;
}

function parseFontScheme(fontScheme: XmlNode | undefined): SourceThemeFontScheme | undefined {
  if (!fontScheme) return undefined;
  const majorLatin = getAttr(getChild(getChild(fontScheme, "majorFont"), "latin"), "typeface");
  const minorLatin = getAttr(getChild(getChild(fontScheme, "minorFont"), "latin"), "typeface");
  const scheme: SourceThemeFontScheme = {
    ...(isNonEmpty(majorLatin) ? { majorLatin } : {}),
    ...(isNonEmpty(minorLatin) ? { minorLatin } : {}),
  };
  return Object.keys(scheme).length > 0 ? scheme : undefined;
}

function booleanAttr(node: XmlNode | undefined, name: string): boolean | undefined {
  const value = getAttr(node, name);
  if (value === undefined) return undefined;
  return isTrue(value);
}

function isNonEmpty(value: string | undefined): value is string {
  return value !== undefined && value !== "";
}
