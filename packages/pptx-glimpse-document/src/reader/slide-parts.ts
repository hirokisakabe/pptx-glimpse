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
  SourceThemeFormatScheme,
} from "../source/index.js";
import {
  asEmu,
  asOoxmlAngle,
  type SourceColor,
  type SourceEffectList,
  type SourceFill,
  type SourceMasterTextStyles,
} from "../source/index.js";
import { isTrue, numericAttr, parseColorElement, parseFill, parseLine } from "./drawing.js";
import { makeSidecar } from "./raw-node.js";
import { parseShapeTree } from "./shape-tree.js";
import { parseTextStyle } from "./text.js";
import {
  getAttr,
  getAttrs,
  getChild,
  getChildArray,
  localName,
  navigateOrdered,
  type XmlNode,
  type XmlOrderedNode,
} from "./xml.js";

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

const FILL_LOCAL_NAMES: ReadonlySet<string> = new Set([
  "solidFill",
  "gradFill",
  "pattFill",
  "blipFill",
  "noFill",
]);

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
  const txStyles = parseMasterTextStyles(getChild(root, "txStyles"));

  return {
    partPath,
    ...(themePartPath !== undefined ? { themePartPath } : {}),
    layoutPartPaths,
    ...(background !== undefined ? { background } : {}),
    ...(colorMap !== undefined ? { colorMap } : {}),
    ...(txStyles !== undefined ? { txStyles } : {}),
    shapes: parseShapeTree(getChild(cSld, "spTree"), partPath, nextId, orderedSpTree),
    handle: { partPath },
  };
}

function parseMasterTextStyles(txStyles: XmlNode | undefined): SourceMasterTextStyles | undefined {
  if (txStyles === undefined) return undefined;
  const titleStyle = parseTextStyle(getChild(txStyles, "titleStyle"));
  const bodyStyle = parseTextStyle(getChild(txStyles, "bodyStyle"));
  const otherStyle = parseTextStyle(getChild(txStyles, "otherStyle"));
  const parsed: SourceMasterTextStyles = {
    ...(titleStyle !== undefined ? { titleStyle } : {}),
    ...(bodyStyle !== undefined ? { bodyStyle } : {}),
    ...(otherStyle !== undefined ? { otherStyle } : {}),
  };
  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

/** parse 済み theme root (`a:theme`) から `SourceTheme` を組み立てる。 */
export function parseTheme(
  root: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderedRoot?: readonly XmlOrderedNode[],
): SourceTheme {
  const name = getAttr(root, "name");
  const themeElements = getChild(root, "themeElements");
  const colorScheme = parseColorScheme(getChild(themeElements, "clrScheme"));
  const fontScheme = parseFontScheme(getChild(themeElements, "fontScheme"));
  const formatScheme = parseFormatScheme(getChild(themeElements, "fmtScheme"), nextId, orderedRoot);

  return {
    partPath,
    ...(name !== undefined ? { name } : {}),
    ...(colorScheme !== undefined ? { colorScheme } : {}),
    ...(fontScheme !== undefined ? { fontScheme } : {}),
    ...(formatScheme !== undefined ? { formatScheme } : {}),
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
  const majorFont = getChild(fontScheme, "majorFont");
  const minorFont = getChild(fontScheme, "minorFont");
  const majorLatin = getAttr(getChild(majorFont, "latin"), "typeface");
  const minorLatin = getAttr(getChild(minorFont, "latin"), "typeface");
  const majorEastAsian = resolveEastAsianFont(majorFont);
  const minorEastAsian = resolveEastAsianFont(minorFont);
  const majorComplexScript = getAttr(getChild(majorFont, "cs"), "typeface");
  const minorComplexScript = getAttr(getChild(minorFont, "cs"), "typeface");
  const majorJapanese = findScriptFont(majorFont, "Jpan");
  const minorJapanese = findScriptFont(minorFont, "Jpan");
  const scheme: SourceThemeFontScheme = {
    ...(isNonEmpty(majorLatin) ? { majorLatin } : {}),
    ...(isNonEmpty(minorLatin) ? { minorLatin } : {}),
    ...(isNonEmpty(majorEastAsian) ? { majorEastAsian } : {}),
    ...(isNonEmpty(minorEastAsian) ? { minorEastAsian } : {}),
    // Preserve explicit empty complex script fonts; parser path resolves +mn-cs to "" for them.
    ...(majorComplexScript !== undefined ? { majorComplexScript } : {}),
    ...(minorComplexScript !== undefined ? { minorComplexScript } : {}),
    ...(isNonEmpty(majorJapanese) ? { majorJapanese } : {}),
    ...(isNonEmpty(minorJapanese) ? { minorJapanese } : {}),
  };
  return Object.keys(scheme).length > 0 ? scheme : undefined;
}

function resolveEastAsianFont(fontNode: XmlNode | undefined): string | undefined {
  const typeface = getAttr(getChild(fontNode, "ea"), "typeface");
  return isNonEmpty(typeface) ? typeface : findScriptFont(fontNode, "Jpan");
}

function findScriptFont(fontNode: XmlNode | undefined, script: string): string | undefined {
  const font = getChildArray(fontNode, "font").find(
    (candidate) => getAttr(candidate, "script") === script,
  );
  return getAttr(font, "typeface");
}

function parseFormatScheme(
  fmtScheme: XmlNode | undefined,
  nextId: () => RawSidecarId,
  orderedRoot?: readonly XmlOrderedNode[],
): SourceThemeFormatScheme | undefined {
  if (fmtScheme === undefined) return undefined;
  const orderedFmtScheme =
    orderedRoot !== undefined ? navigateOrdered(orderedRoot, ["themeElements", "fmtScheme"]) : [];
  const fillStyles = parseFillStyleList(
    getChild(fmtScheme, "fillStyleLst"),
    orderedFmtScheme,
    "fillStyleLst",
    nextId,
  );
  const backgroundFillStyles = parseFillStyleList(
    getChild(fmtScheme, "bgFillStyleLst"),
    orderedFmtScheme,
    "bgFillStyleLst",
    nextId,
  );
  const lineStyles = getChildArray(getChild(fmtScheme, "lnStyleLst"), "ln").flatMap((ln) => {
    const line = parseLine(ln, nextId);
    return line !== undefined ? [line] : [];
  });
  const effectStyles = getChildArray(getChild(fmtScheme, "effectStyleLst"), "effectStyle").map(
    (effectStyle) => parseEffectList(getChild(effectStyle, "effectLst")),
  );

  if (
    fillStyles.length === 0 &&
    backgroundFillStyles.length === 0 &&
    lineStyles.length === 0 &&
    effectStyles.length === 0
  ) {
    return undefined;
  }

  return { fillStyles, lineStyles, effectStyles, backgroundFillStyles };
}

function parseFillStyleList(
  list: XmlNode | undefined,
  orderedFmtScheme: readonly XmlOrderedNode[] | undefined,
  listName: string,
  nextId: () => RawSidecarId,
): SourceFill[] {
  if (list === undefined) return [];
  const orderedList = orderedFmtScheme?.find((node) => listName in node)?.[listName];
  if (!Array.isArray(orderedList)) return [];
  const orderedItems = orderedList as readonly XmlOrderedNode[];

  const fills: SourceFill[] = [];
  const tagCounters: Record<string, number> = {};
  for (const orderedChild of orderedItems) {
    const qualifiedName = Object.keys(orderedChild).find((key) => FILL_LOCAL_NAMES.has(key));
    if (qualifiedName === undefined) continue;
    const name = localName(qualifiedName);
    const index = tagCounters[name] ?? 0;
    tagCounters[name] = index + 1;
    const values = getChildArray(list, name);
    const value = values[index];
    if (value === undefined) continue;
    const fill = parseFill({ [qualifiedName]: value }, nextId);
    if (fill !== undefined) fills.push(fill);
  }
  return fills;
}

function parseEffectList(effectList: XmlNode | undefined): SourceEffectList | undefined {
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

function parseOuterShadow(node: XmlNode | undefined): SourceEffectList["outerShadow"] | undefined {
  const color = parseColorElement(node);
  if (node === undefined || color === undefined) return undefined;
  return {
    blurRadius: asEmu(numericAttr(node, "blurRad") ?? 0),
    distance: asEmu(numericAttr(node, "dist") ?? 0),
    direction: asOoxmlAngle(numericAttr(node, "dir") ?? 0),
    color,
    alignment: getAttr(node, "algn") ?? "b",
    rotateWithShape: getAttr(node, "rotWithShape") !== "0",
  };
}

function parseInnerShadow(node: XmlNode | undefined): SourceEffectList["innerShadow"] | undefined {
  const color = parseColorElement(node);
  if (node === undefined || color === undefined) return undefined;
  return {
    blurRadius: asEmu(numericAttr(node, "blurRad") ?? 0),
    distance: asEmu(numericAttr(node, "dist") ?? 0),
    direction: asOoxmlAngle(numericAttr(node, "dir") ?? 0),
    color,
  };
}

function parseGlow(node: XmlNode | undefined): SourceEffectList["glow"] | undefined {
  const color = parseColorElement(node);
  if (node === undefined || color === undefined) return undefined;
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

function booleanAttr(node: XmlNode | undefined, name: string): boolean | undefined {
  const value = getAttr(node, name);
  if (value === undefined) return undefined;
  return isTrue(value);
}

function isNonEmpty(value: string | undefined): value is string {
  return value !== undefined && value !== "";
}
