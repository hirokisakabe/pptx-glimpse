import type { Slide, Background } from "../model/slide.js";
import type {
  SlideElement,
  ShapeElement,
  ConnectorElement,
  GroupElement,
  Transform,
  Geometry,
} from "../model/shape.js";
import type { ImageElement } from "../model/image.js";
import type { ChartElement } from "../model/chart.js";
import type { TableElement } from "../model/table.js";
import type {
  TextBody,
  BodyProperties,
  Paragraph,
  TextRun,
  RunProperties,
  Hyperlink,
  BulletType,
  AutoNumScheme,
  DefaultTextStyle,
  DefaultRunProperties,
} from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseEffectList } from "./effect-parser.js";
import { parseChart } from "./chart-parser.js";
import { parseCustomGeometry } from "./custom-geometry-parser.js";
import { parseTable } from "./table-parser.js";
import {
  parseRelationships,
  resolveRelationshipTarget,
  buildRelsPath,
} from "./relationship-parser.js";
import { hundredthPointToPoint } from "../utils/emu.js";
import { parseListStyle, parseDefaultRunProperties } from "./text-style-parser.js";

const WARN_PREFIX = "[pptx-glimpse]";

export function parseSlide(
  slideXml: string,
  slidePath: string,
  slideNumber: number,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): Slide {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(slideXml) as any;
  const sld = parsed.sld;
  if (!sld) {
    console.warn(`${WARN_PREFIX} Slide ${slideNumber}: missing root element "sld" in XML`);
  }

  const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const fillContext: FillParseContext = { rels, archive, basePath: slidePath };
  const background = parseBackground(sld?.cSld?.bg, colorResolver, fillContext);
  const elements = parseShapeTree(
    sld?.cSld?.spTree,
    rels,
    slidePath,
    archive,
    colorResolver,
    `Slide ${slideNumber}`,
    fillContext,
    fontScheme,
  );

  return { slideNumber, background, elements };
}

function parseBackground(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  bgNode: any,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  if (!bgNode) return null;

  const bgPr = bgNode.bgPr;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mergeChildElements(spTree: any, source: any): void {
  const tags = ["sp", "pic", "cxnSp", "grpSp", "graphicFrame"];
  for (const tag of tags) {
    const items = source[tag];
    if (!items) continue;
    if (!spTree[tag]) {
      spTree[tag] = [];
    }
    const arr = Array.isArray(items) ? items : [items];
    for (const item of arr) {
      spTree[tag].push(item);
    }
  }
}

export function parseShapeTree(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  spTree: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  context?: string,
  fillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
): SlideElement[] {
  if (!spTree) return [];

  // mc:AlternateContent の処理: Choice 内の要素を spTree にマージ
  const alternateContents = spTree.AlternateContent ?? [];
  for (const ac of alternateContents) {
    const choices = Array.isArray(ac.Choice) ? ac.Choice : ac.Choice ? [ac.Choice] : [];
    for (const choice of choices) {
      mergeChildElements(spTree, choice);
    }
  }

  const ctx = context ?? slidePath;
  const elements: SlideElement[] = [];

  const shapes = spTree.sp ?? [];
  for (const sp of shapes) {
    const shape = parseShape(sp, colorResolver, rels, fillContext, fontScheme);
    if (shape) {
      elements.push(shape);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: shape skipped (parse returned null)`);
    }
  }

  const pics = spTree.pic ?? [];
  for (const pic of pics) {
    const img = parseImage(pic, rels, slidePath, archive, colorResolver);
    if (img) {
      elements.push(img);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: image skipped (parse returned null)`);
    }
  }

  const cxnSps = spTree.cxnSp ?? [];
  for (const cxn of cxnSps) {
    const connector = parseConnector(cxn, colorResolver);
    if (connector) {
      elements.push(connector);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: connector skipped (parse returned null)`);
    }
  }

  const grpSps = spTree.grpSp ?? [];
  for (const grp of grpSps) {
    const group = parseGroup(grp, rels, slidePath, archive, colorResolver, fillContext, fontScheme);
    if (group) {
      elements.push(group);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: group skipped (parse returned null)`);
    }
  }

  const graphicFrames = spTree.graphicFrame ?? [];
  for (const gf of graphicFrames) {
    const chart = parseGraphicFrame(gf, rels, slidePath, archive, colorResolver, fontScheme);
    if (chart) {
      elements.push(chart);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: graphicFrame skipped (parse returned null)`);
    }
  }

  return elements;
}

function parseShape(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sp: any,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
): ShapeElement | null {
  const spPr = sp.spPr;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm);
  if (!transform) return null;

  const geometry = parseGeometry(spPr);
  const fill = parseFillFromNode(spPr, colorResolver, fillContext);
  const outline = parseOutline(spPr.ln, colorResolver);
  const textBody = parseTextBody(sp.txBody, colorResolver, rels, fontScheme);
  const effects = parseEffectList(spPr.effectLst, colorResolver);

  const ph = sp.nvSpPr?.nvPr?.ph;
  const placeholderType = ph ? (ph["@_type"] ?? "body") : undefined;
  const placeholderIdx = ph?.["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;

  return {
    type: "shape",
    transform,
    geometry,
    fill,
    outline,
    textBody,
    effects,
    ...(placeholderType !== undefined && { placeholderType }),
    ...(placeholderIdx !== undefined && { placeholderIdx }),
  };
}

function parseImage(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pic: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): ImageElement | null {
  const spPr = pic.spPr;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm);
  if (!transform) return null;

  const blipFill = pic.blipFill;
  const rId = blipFill?.blip?.["@_r:embed"] ?? blipFill?.blip?.["@_embed"];
  if (!rId) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const mediaPath = resolveRelationshipTarget(slidePath, rel.target);
  const mediaData = archive.media.get(mediaPath);
  if (!mediaData) return null;

  const ext = mediaPath.split(".").pop()?.toLowerCase() ?? "png";
  const mimeMap: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml",
    emf: "image/emf",
    wmf: "image/wmf",
  };
  const mimeType = mimeMap[ext] ?? "image/png";
  const imageData = mediaData.toString("base64");
  const effects = parseEffectList(spPr.effectLst, colorResolver);

  return {
    type: "image",
    transform,
    imageData,
    mimeType,
    effects,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseConnector(cxn: any, colorResolver: ColorResolver): ConnectorElement | null {
  const spPr = cxn.spPr;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm);
  if (!transform) return null;

  const outline = parseOutline(spPr.ln, colorResolver);
  const effects = parseEffectList(spPr.effectLst, colorResolver);

  return { type: "connector", transform, outline, effects };
}

function parseGroup(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  grp: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  parentFillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
): GroupElement | null {
  const grpSpPr = grp.grpSpPr;
  if (!grpSpPr) return null;

  const transform = parseTransform(grpSpPr.xfrm);
  if (!transform) return null;

  const childOff = grpSpPr.xfrm?.chOff;
  const childExt = grpSpPr.xfrm?.chExt;
  const childTransform: Transform = {
    offsetX: Number(childOff?.["@_x"] ?? 0),
    offsetY: Number(childOff?.["@_y"] ?? 0),
    extentWidth: Number(childExt?.["@_cx"] ?? transform.extentWidth),
    extentHeight: Number(childExt?.["@_cy"] ?? transform.extentHeight),
    rotation: 0,
    flipH: false,
    flipV: false,
  };

  const groupFill = parseFillFromNode(grpSpPr, colorResolver, parentFillContext);
  const childFillContext: FillParseContext = {
    rels: parentFillContext?.rels ?? rels,
    archive: parentFillContext?.archive ?? { files: new Map(), media: new Map() },
    basePath: parentFillContext?.basePath ?? slidePath,
    ...(groupFill ? { groupFill } : {}),
  };

  const children = parseShapeTree(
    grp,
    rels,
    slidePath,
    archive,
    colorResolver,
    undefined,
    childFillContext,
    fontScheme,
  );
  const effects = parseEffectList(grpSpPr.effectLst, colorResolver);

  return { type: "group", transform, childTransform, children, effects };
}

function parseGraphicFrame(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  gf: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): ChartElement | TableElement | GroupElement | null {
  const xfrm = gf.xfrm;
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const graphicData = gf.graphic?.graphicData;
  if (!graphicData) return null;

  // Chart
  const chartRef = graphicData.chart;
  if (chartRef) {
    const rId = chartRef["@_r:id"] ?? chartRef["@_id"];
    if (!rId) return null;

    const rel = rels.get(rId);
    if (!rel) return null;

    const chartPath = resolveRelationshipTarget(slidePath, rel.target);
    const chartXml = archive.files.get(chartPath);
    if (!chartXml) return null;

    const chartData = parseChart(chartXml, colorResolver);
    if (!chartData) return null;

    return { type: "chart", transform, chart: chartData };
  }

  // Table
  const tblNode = graphicData.tbl;
  if (tblNode) {
    const tableData = parseTable(tblNode, colorResolver, fontScheme);
    if (!tableData) return null;

    return { type: "table", transform, table: tableData };
  }

  // SmartArt (Diagram)
  if (graphicData["@_uri"] === "http://schemas.openxmlformats.org/drawingml/2006/diagram") {
    return parseSmartArt(
      graphicData,
      transform,
      rels,
      slidePath,
      archive,
      colorResolver,
      fontScheme,
    );
  }

  return null;
}

function parseSmartArt(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  graphicData: any,
  transform: Transform,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): GroupElement | null {
  // dgm:relIds → removeNSPrefix → relIds
  const relIds = graphicData.relIds;
  if (!relIds) return null;

  // r:dm 属性 (data model relationship ID)
  const dmRId = relIds["@_r:dm"] ?? relIds["@_dm"];
  if (!dmRId) return null;

  const dmRel = rels.get(dmRId);
  if (!dmRel) return null;

  const dataPath = resolveRelationshipTarget(slidePath, dmRel.target);

  // data の .rels から diagramDrawing リレーションシップを探す
  const dataRelsPath = buildRelsPath(dataPath);
  const dataRelsXml = archive.files.get(dataRelsPath);

  let drawingPath: string | null = null;

  if (dataRelsXml) {
    const dataRels = parseRelationships(dataRelsXml);
    for (const [, rel] of dataRels) {
      if (rel.type.includes("diagramDrawing")) {
        drawingPath = resolveRelationshipTarget(dataPath, rel.target);
        break;
      }
    }
  }

  // フォールバック: slide rels から diagramDrawing を探す
  if (!drawingPath) {
    for (const [, rel] of rels) {
      if (rel.type.includes("diagramDrawing")) {
        drawingPath = resolveRelationshipTarget(slidePath, rel.target);
        break;
      }
    }
  }

  if (!drawingPath) return null;

  const drawingXml = archive.files.get(drawingPath);
  if (!drawingXml) return null;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(drawingXml) as any;
  const spTree = parsed.drawing?.spTree;
  if (!spTree) return null;

  // drawing XML 用のリレーションシップを取得
  const drawingRelsPath = buildRelsPath(drawingPath);
  const drawingRelsXml = archive.files.get(drawingRelsPath);
  const drawingRels = drawingRelsXml
    ? parseRelationships(drawingRelsXml)
    : new Map<string, Relationship>();

  // grpSpPr から childTransform を取得
  const grpXfrm = spTree.grpSpPr?.xfrm;
  const childOff = grpXfrm?.chOff;
  const childExt = grpXfrm?.chExt;
  const childTransform: Transform = {
    offsetX: Number(childOff?.["@_x"] ?? 0),
    offsetY: Number(childOff?.["@_y"] ?? 0),
    extentWidth: Number(childExt?.["@_cx"] ?? transform.extentWidth),
    extentHeight: Number(childExt?.["@_cy"] ?? transform.extentHeight),
    rotation: 0,
    flipH: false,
    flipV: false,
  };

  const fillContext: FillParseContext = {
    rels: drawingRels,
    archive,
    basePath: drawingPath,
  };

  const children = parseShapeTree(
    spTree,
    drawingRels,
    drawingPath,
    archive,
    colorResolver,
    undefined,
    fillContext,
    fontScheme,
  );

  if (children.length === 0) return null;

  return {
    type: "group",
    transform,
    childTransform,
    children,
    effects: null,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseTransform(xfrm: any): Transform | null {
  if (!xfrm) return null;

  const off = xfrm.off;
  const ext = xfrm.ext;
  if (!off || !ext) return null;

  let offsetX = Number(off["@_x"] ?? 0);
  let offsetY = Number(off["@_y"] ?? 0);
  let extentWidth = Number(ext["@_cx"] ?? 0);
  let extentHeight = Number(ext["@_cy"] ?? 0);
  let rotation = Number(xfrm["@_rot"] ?? 0);

  if (Number.isNaN(offsetX)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform offsetX, defaulting to 0`);
    offsetX = 0;
  }
  if (Number.isNaN(offsetY)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform offsetY, defaulting to 0`);
    offsetY = 0;
  }
  if (Number.isNaN(extentWidth)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform extentWidth, defaulting to 0`);
    extentWidth = 0;
  }
  if (Number.isNaN(extentHeight)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform extentHeight, defaulting to 0`);
    extentHeight = 0;
  }
  if (Number.isNaN(rotation)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform rotation, defaulting to 0`);
    rotation = 0;
  }

  return {
    offsetX,
    offsetY,
    extentWidth,
    extentHeight,
    rotation: rotation / 60000,
    flipH: xfrm["@_flipH"] === "1" || xfrm["@_flipH"] === "true",
    flipV: xfrm["@_flipV"] === "1" || xfrm["@_flipV"] === "true",
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseGeometry(spPr: any): Geometry {
  if (spPr.prstGeom) {
    const preset = spPr.prstGeom["@_prst"] ?? "rect";
    const avLst = spPr.prstGeom.avLst;
    const adjustValues: Record<string, number> = {};

    if (avLst?.gd) {
      const guides = Array.isArray(avLst.gd) ? avLst.gd : [avLst.gd];
      for (const gd of guides) {
        const name = gd["@_name"] as string;
        const fmla = gd["@_fmla"] as string;
        const match = fmla?.match(/val\s+(\d+)/);
        if (name && match) {
          adjustValues[name] = Number(match[1]);
        }
      }
    }

    return { type: "preset", preset, adjustValues };
  }

  if (spPr.custGeom) {
    const paths = parseCustomGeometry(spPr.custGeom);
    if (paths) {
      return { type: "custom", paths };
    }
  }

  return { type: "preset", preset: "rect", adjustValues: {} };
}

export function parseTextBody(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  txBody: any,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
): TextBody | null {
  if (!txBody) return null;

  const bodyPr = txBody.bodyPr;

  let autoFit: BodyProperties["autoFit"] = "noAutofit";
  let fontScale = 1;
  let lnSpcReduction = 0;
  if (bodyPr?.normAutofit !== undefined) {
    autoFit = "normAutofit";
    const normAutofit = bodyPr.normAutofit;
    if (typeof normAutofit === "object" && normAutofit !== null) {
      fontScale = normAutofit["@_fontScale"] ? Number(normAutofit["@_fontScale"]) / 100000 : 1;
      lnSpcReduction = normAutofit["@_lnSpcReduction"]
        ? Number(normAutofit["@_lnSpcReduction"]) / 100000
        : 0;
    }
  } else if (bodyPr?.spAutoFit !== undefined) {
    autoFit = "spAutofit";
  }

  const bodyProperties: BodyProperties = {
    anchor: (bodyPr?.["@_anchor"] as "t" | "ctr" | "b") ?? "t",
    marginLeft: Number(bodyPr?.["@_lIns"] ?? 91440),
    marginRight: Number(bodyPr?.["@_rIns"] ?? 91440),
    marginTop: Number(bodyPr?.["@_tIns"] ?? 45720),
    marginBottom: Number(bodyPr?.["@_bIns"] ?? 45720),
    wrap: (bodyPr?.["@_wrap"] as "square" | "none") ?? "square",
    autoFit,
    fontScale,
    lnSpcReduction,
  };

  const lstStyle = parseListStyle(txBody.lstStyle);

  const paragraphs: Paragraph[] = [];
  const pList = txBody.p ?? [];
  for (const p of pList) {
    paragraphs.push(parseParagraph(p, colorResolver, rels, fontScheme, lstStyle));
  }

  if (paragraphs.length === 0) return null;

  return { paragraphs, bodyProperties };
}

const VALID_AUTO_NUM_SCHEMES = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseBullet(pPr: any, colorResolver: ColorResolver) {
  let bullet: BulletType | null = null;
  let bulletFont: string | null = null;
  let bulletColor = colorResolver.resolve(pPr?.buClr);
  if (!pPr?.buClr) bulletColor = null;
  const bulletSizePct: number | null = pPr?.buSzPct ? Number(pPr.buSzPct["@_val"]) : null;

  if (pPr?.buNone !== undefined) {
    bullet = { type: "none" };
  } else if (pPr?.buChar) {
    bullet = { type: "char", char: pPr.buChar["@_char"] ?? "\u2022" };
  } else if (pPr?.buAutoNum) {
    const scheme = pPr.buAutoNum["@_type"] ?? "arabicPeriod";
    bullet = {
      type: "autoNum",
      scheme: VALID_AUTO_NUM_SCHEMES.has(scheme) ? (scheme as AutoNumScheme) : "arabicPeriod",
      startAt: Number(pPr.buAutoNum["@_startAt"] ?? 1),
    };
  }

  if (pPr?.buFont) {
    bulletFont = pPr.buFont["@_typeface"] ?? null;
  }

  return { bullet, bulletFont, bulletColor, bulletSizePct };
}

function parseParagraph(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  p: any,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
  lstStyle?: DefaultTextStyle,
): Paragraph {
  const pPr = p.pPr;
  const level = Number(pPr?.["@_lvl"] ?? 0);

  // lstStyle からレベル対応のデフォルト段落プロパティを取得
  const lstLevelProps = lstStyle?.levels[level];

  const { bullet, bulletFont, bulletColor, bulletSizePct } = parseBullet(pPr, colorResolver);
  const properties = {
    alignment: (pPr?.["@_algn"] as "l" | "ctr" | "r" | "just") ?? lstLevelProps?.alignment ?? "l",
    lineSpacing: pPr?.lnSpc?.spcPct ? Number(pPr.lnSpc.spcPct["@_val"]) : null,
    spaceBefore: pPr?.spcBef?.spcPts ? Number(pPr.spcBef.spcPts["@_val"]) : 0,
    spaceAfter: pPr?.spcAft?.spcPts ? Number(pPr.spcAft.spcPts["@_val"]) : 0,
    level,
    bullet,
    bulletFont,
    bulletColor,
    bulletSizePct,
    marginLeft:
      pPr?.["@_marL"] !== undefined ? Number(pPr["@_marL"]) : (lstLevelProps?.marginLeft ?? 0),
    indent:
      pPr?.["@_indent"] !== undefined ? Number(pPr["@_indent"]) : (lstLevelProps?.indent ?? 0),
  };

  // defRPr のマージ: pPr.defRPr > lstStyle.lvl.defRPr
  const pPrDefRPr = parseDefaultRunProperties(pPr?.defRPr);
  const lstDefRPr = lstLevelProps?.defaultRunProperties;
  const mergedDefaults = mergeDefaultRunProperties(pPrDefRPr, lstDefRPr);

  const runs: TextRun[] = [];
  const rList = p.r ?? [];
  for (const r of rList) {
    const text = r.t ?? "";
    const textContent = typeof text === "object" ? (text["#text"] ?? "") : String(text);
    const rPr = r.rPr;
    const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
    runs.push({ text: textContent, properties: runProps });
  }

  return { runs, properties };
}

function mergeDefaultRunProperties(
  primary?: DefaultRunProperties,
  secondary?: DefaultRunProperties,
): DefaultRunProperties | undefined {
  if (!primary && !secondary) return undefined;
  if (!primary) return secondary;
  if (!secondary) return primary;
  return {
    fontSize: primary.fontSize ?? secondary.fontSize,
    fontFamily: primary.fontFamily ?? secondary.fontFamily,
    fontFamilyEa: primary.fontFamilyEa ?? secondary.fontFamilyEa,
    bold: primary.bold ?? secondary.bold,
    italic: primary.italic ?? secondary.italic,
    underline: primary.underline ?? secondary.underline,
    strikethrough: primary.strikethrough ?? secondary.strikethrough,
  };
}

function resolveThemeFont(typeface: string | null, fontScheme?: FontScheme | null): string | null {
  if (!typeface || !fontScheme) return typeface;
  switch (typeface) {
    case "+mj-lt":
      return fontScheme.majorFont;
    case "+mn-lt":
      return fontScheme.minorFont;
    case "+mj-ea":
      return fontScheme.majorFontEa;
    case "+mn-ea":
      return fontScheme.minorFontEa;
    default:
      return typeface;
  }
}

function parseRunProperties(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  rPr: any,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
  defaults?: DefaultRunProperties,
): RunProperties {
  if (!rPr) {
    return {
      fontSize: defaults?.fontSize ?? null,
      fontFamily: resolveThemeFont(defaults?.fontFamily ?? null, fontScheme),
      fontFamilyEa: resolveThemeFont(defaults?.fontFamilyEa ?? null, fontScheme),
      bold: defaults?.bold ?? false,
      italic: defaults?.italic ?? false,
      underline: defaults?.underline ?? false,
      strikethrough: defaults?.strikethrough ?? false,
      color: null,
      baseline: 0,
      hyperlink: null,
    };
  }

  const fontSize = rPr["@_sz"]
    ? hundredthPointToPoint(Number(rPr["@_sz"]))
    : (defaults?.fontSize ?? null);
  const fontFamily = resolveThemeFont(
    rPr.latin?.["@_typeface"] ?? defaults?.fontFamily ?? null,
    fontScheme,
  );
  const fontFamilyEa = resolveThemeFont(
    rPr.ea?.["@_typeface"] ?? defaults?.fontFamilyEa ?? null,
    fontScheme,
  );
  const bold =
    rPr["@_b"] !== undefined
      ? rPr["@_b"] === "1" || rPr["@_b"] === "true"
      : (defaults?.bold ?? false);
  const italic =
    rPr["@_i"] !== undefined
      ? rPr["@_i"] === "1" || rPr["@_i"] === "true"
      : (defaults?.italic ?? false);
  const underline =
    rPr["@_u"] !== undefined ? rPr["@_u"] !== "none" : (defaults?.underline ?? false);
  const strikethrough =
    rPr["@_strike"] !== undefined
      ? rPr["@_strike"] !== "noStrike"
      : (defaults?.strikethrough ?? false);
  const baseline = rPr["@_baseline"] ? Number(rPr["@_baseline"]) / 1000 : 0;

  let color = colorResolver.resolve(rPr.solidFill ?? rPr);
  if (!rPr.solidFill && !rPr.srgbClr && !rPr.schemeClr && !rPr.sysClr) {
    color = null;
  }

  const hyperlink = parseHyperlink(rPr.hlinkClick, rels);

  return {
    fontSize,
    fontFamily,
    fontFamilyEa,
    bold,
    italic,
    underline,
    strikethrough,
    color,
    baseline,
    hyperlink,
  };
}

function parseHyperlink(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  hlinkClick: any,
  rels?: Map<string, Relationship>,
): Hyperlink | null {
  if (!hlinkClick) return null;

  const rId = hlinkClick["@_r:id"] ?? hlinkClick["@_id"];
  if (!rId || !rels) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const tooltip = hlinkClick["@_tooltip"] as string | undefined;
  return { url: rel.target, ...(tooltip && { tooltip }) };
}
