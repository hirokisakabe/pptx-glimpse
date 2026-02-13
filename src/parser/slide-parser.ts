import type { Slide, Background } from "../model/slide.js";
import type {
  SlideElement,
  ShapeElement,
  ConnectorElement,
  GroupElement,
  Transform,
  Geometry,
} from "../model/shape.js";
import type { ImageElement, SrcRect, StretchFillRect, TileInfo } from "../model/image.js";
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
  SpacingValue,
  TabStop,
} from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme, FormatScheme } from "../model/theme.js";
import { parseXml, parseXmlOrdered } from "./xml-parser.js";
import type { XmlNode, XmlOrderedNode } from "./xml-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseEffectList } from "./effect-parser.js";
import { resolveShapeStyle } from "./style-reference-resolver.js";
import { parseBlipEffects } from "./blip-effect-parser.js";
import { parseChart } from "./chart-parser.js";
import { parseCustomGeometry } from "./custom-geometry-parser.js";
import { parseTable } from "./table-parser.js";
import {
  parseRelationships,
  resolveRelationshipTarget,
  buildRelsPath,
} from "./relationship-parser.js";
import { hundredthPointToPoint } from "../utils/emu.js";
import {
  parseListStyle,
  parseDefaultRunProperties,
  resolveThemeFont,
} from "./text-style-parser.js";
import { uint8ArrayToBase64 } from "../utils/base64.js";
import { warn, debug } from "../warning-logger.js";

const SHAPE_TAGS = new Set(["sp", "pic", "cxnSp", "grpSp", "graphicFrame"]);

// preserveOrder パース結果を path で辿り、指定ノードの子配列を返す。
export function navigateOrdered(
  ordered: XmlOrderedNode[],
  path: string[],
): XmlOrderedNode[] | null {
  let current: XmlOrderedNode[] = ordered;
  for (const key of path) {
    const entry = current.find((item: XmlOrderedNode) => key in item);
    if (!entry) return null;
    current = entry[key] as XmlOrderedNode[];
    if (!Array.isArray(current)) return null;
  }
  return current;
}

export function parseSlide(
  slideXml: string,
  slidePath: string,
  slideNumber: number,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
  fmtScheme?: FormatScheme,
): Slide {
  const parsed = parseXml(slideXml);
  const sld = parsed.sld as XmlNode | undefined;
  if (!sld) {
    debug("slide.missing", `missing root element "sld" in XML`, `Slide ${slideNumber}`);
  }

  const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const fillContext: FillParseContext = { rels, archive, basePath: slidePath };
  const cSld = (sld?.cSld as XmlNode | undefined) ?? undefined;
  const background = parseBackground(cSld?.bg as XmlNode | undefined, colorResolver, fillContext);

  // ordered パーサーで子要素の出現順序を取得（Z-order 保持のため）
  const orderedParsed = parseXmlOrdered(slideXml);
  const orderedSpTree = navigateOrdered(orderedParsed, ["sld", "cSld", "spTree"]);

  const elements = parseShapeTree(
    cSld?.spTree as XmlNode | undefined,
    rels,
    slidePath,
    archive,
    colorResolver,
    `Slide ${slideNumber}`,
    fillContext,
    fontScheme,
    orderedSpTree,
    fmtScheme,
  );

  const showMasterSpAttr = sld?.["@_showMasterSp"];
  const showMasterSp = showMasterSpAttr !== "0" && showMasterSpAttr !== "false";

  return { slideNumber, background, elements, showMasterSp };
}

function parseBackground(
  bgNode: XmlNode | undefined,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  if (!bgNode) return null;

  const bgPr = bgNode.bgPr as XmlNode | undefined;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

function mergeChildElements(spTree: XmlNode, source: XmlNode): void {
  const tags = ["sp", "pic", "cxnSp", "grpSp", "graphicFrame"];
  for (const tag of tags) {
    const items = source[tag];
    if (!items) continue;
    if (!spTree[tag]) {
      spTree[tag] = [];
    }
    const arr = Array.isArray(items) ? items : [items];
    for (const item of arr) {
      (spTree[tag] as unknown[]).push(item);
    }
  }
}

export function parseShapeTree(
  spTree: XmlNode | undefined,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  context?: string,
  fillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
  orderedChildren?: XmlOrderedNode[] | null,
  fmtScheme?: FormatScheme,
): SlideElement[] {
  if (!spTree) return [];

  // orderedChildren が提供されている場合はドキュメント順でイテレート
  if (orderedChildren) {
    return parseShapeTreeOrdered(
      spTree,
      orderedChildren,
      rels,
      slidePath,
      archive,
      colorResolver,
      context,
      fillContext,
      fontScheme,
      fmtScheme,
    );
  }

  // フォールバック: タイプ別イテレーション（後方互換）
  // mc:AlternateContent の処理: Choice 内の要素を spTree にマージ
  const alternateContents = (spTree.AlternateContent as XmlNode[] | undefined) ?? [];
  for (const ac of alternateContents) {
    const choices = Array.isArray(ac.Choice) ? ac.Choice : ac.Choice ? [ac.Choice] : [];
    for (const choice of choices) {
      mergeChildElements(spTree, choice as XmlNode);
    }
  }

  const ctx = context ?? slidePath;
  const elements: SlideElement[] = [];

  const shapes = (spTree.sp as XmlNode[] | undefined) ?? [];
  for (const sp of shapes) {
    const shape = parseShape(
      sp,
      colorResolver,
      rels,
      fillContext,
      fontScheme,
      undefined,
      fmtScheme,
    );
    if (shape) {
      elements.push(shape);
    } else {
      debug("shape.skipped", "shape skipped (parse returned null)", ctx);
    }
  }

  const pics = (spTree.pic as XmlNode[] | undefined) ?? [];
  for (const pic of pics) {
    const img = parseImage(pic, rels, slidePath, archive, colorResolver);
    if (img) {
      elements.push(img);
    } else {
      debug("image.skipped", "image skipped (parse returned null)", ctx);
    }
  }

  const cxnSps = (spTree.cxnSp as XmlNode[] | undefined) ?? [];
  for (const cxn of cxnSps) {
    const connector = parseConnector(cxn, colorResolver, fmtScheme);
    if (connector) {
      elements.push(connector);
    } else {
      debug("connector.skipped", "connector skipped (parse returned null)", ctx);
    }
  }

  const grpSps = (spTree.grpSp as XmlNode[] | undefined) ?? [];
  for (const grp of grpSps) {
    const group = parseGroup(
      grp,
      rels,
      slidePath,
      archive,
      colorResolver,
      fillContext,
      fontScheme,
      undefined,
      fmtScheme,
    );
    if (group) {
      elements.push(group);
    } else {
      debug("group.skipped", "group skipped (parse returned null)", ctx);
    }
  }

  const graphicFrames = (spTree.graphicFrame as XmlNode[] | undefined) ?? [];
  for (const gf of graphicFrames) {
    const chart = parseGraphicFrame(gf, rels, slidePath, archive, colorResolver, fontScheme);
    if (chart) {
      elements.push(chart);
    } else {
      debug("graphicFrame.skipped", "graphicFrame skipped (parse returned null)", ctx);
    }
  }

  return elements;
}

// orderedChildren を使ってドキュメント順で要素をイテレートする
function parseShapeTreeOrdered(
  spTree: XmlNode,
  orderedChildren: XmlOrderedNode[],
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  context?: string,
  fillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
  fmtScheme?: FormatScheme,
): SlideElement[] {
  const ctx = context ?? slidePath;
  const elements: SlideElement[] = [];
  const tagCounters: Record<string, number> = {};

  for (const child of orderedChildren) {
    const tag = Object.keys(child).find((k) => k !== ":@");
    if (!tag) continue;

    if (tag === "AlternateContent") {
      // AlternateContent をインライン処理
      const acIndex = tagCounters["AlternateContent"] ?? 0;
      tagCounters["AlternateContent"] = acIndex + 1;

      const acList = (spTree.AlternateContent as XmlNode[] | undefined) ?? [];
      const acData = acList[acIndex] as XmlNode | undefined;
      if (!acData) continue;

      const choices = Array.isArray(acData.Choice)
        ? acData.Choice
        : acData.Choice
          ? [acData.Choice]
          : [];
      const firstChoice = choices[0] as XmlNode | undefined;
      if (!firstChoice) continue;

      // ordered 結果から Choice の子要素を取得
      const acOrderedChildren = child.AlternateContent as XmlOrderedNode[] | undefined;

      const choiceOrdered = Array.isArray(acOrderedChildren)
        ? acOrderedChildren.find((c: XmlOrderedNode) => "Choice" in c)
        : null;
      const choiceChildren = choiceOrdered?.Choice as XmlOrderedNode[] | undefined;

      if (Array.isArray(choiceChildren)) {
        // Choice 内の子要素をドキュメント順で処理
        const choiceTagCounters: Record<string, number> = {};
        for (const choiceChild of choiceChildren) {
          const choiceTag = Object.keys(choiceChild).find((k) => k !== ":@");
          if (!choiceTag || !SHAPE_TAGS.has(choiceTag)) continue;

          const choiceIdx = choiceTagCounters[choiceTag] ?? 0;
          choiceTagCounters[choiceTag] = choiceIdx + 1;

          const items = firstChoice[choiceTag];
          const arr = Array.isArray(items) ? items : items ? [items] : [];
          const element = arr[choiceIdx] as XmlNode | undefined;
          if (!element) continue;

          parseAndPushElement(
            choiceTag,
            element,
            choiceChild,
            elements,
            rels,
            slidePath,
            archive,
            colorResolver,
            ctx,
            fillContext,
            fontScheme,
            fmtScheme,
          );
        }
      }
      continue;
    }

    if (!SHAPE_TAGS.has(tag)) continue;

    const index = tagCounters[tag] ?? 0;
    tagCounters[tag] = index + 1;

    const tagArray = spTree[tag] as XmlNode[] | undefined;
    const element = tagArray?.[index];
    if (!element) continue;

    parseAndPushElement(
      tag,
      element,
      child,
      elements,
      rels,
      slidePath,
      archive,
      colorResolver,
      ctx,
      fillContext,
      fontScheme,
      fmtScheme,
    );
  }

  return elements;
}

// タグに応じた要素パース処理を行い elements に追加する
function parseAndPushElement(
  tag: string,
  element: XmlNode,
  orderedNode: XmlOrderedNode,
  elements: SlideElement[],
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  ctx: string,
  fillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
  fmtScheme?: FormatScheme,
): void {
  switch (tag) {
    case "sp": {
      const shape = parseShape(
        element,
        colorResolver,
        rels,
        fillContext,
        fontScheme,
        orderedNode,
        fmtScheme,
      );
      if (shape) {
        elements.push(shape);
      } else {
        debug("shape.skipped", "shape skipped (parse returned null)", ctx);
      }
      break;
    }
    case "pic": {
      const img = parseImage(element, rels, slidePath, archive, colorResolver);
      if (img) {
        elements.push(img);
      } else {
        debug("image.skipped", "image skipped (parse returned null)", ctx);
      }
      break;
    }
    case "cxnSp": {
      const connector = parseConnector(element, colorResolver, fmtScheme);
      if (connector) {
        elements.push(connector);
      } else {
        debug("connector.skipped", "connector skipped (parse returned null)", ctx);
      }
      break;
    }
    case "grpSp": {
      // ordered 結果からグループの子要素順序を取得
      const grpOrderedChildren = (orderedNode[tag] as XmlOrderedNode[] | undefined) ?? null;
      const group = parseGroup(
        element,
        rels,
        slidePath,
        archive,
        colorResolver,
        fillContext,
        fontScheme,
        grpOrderedChildren,
        fmtScheme,
      );
      if (group) {
        elements.push(group);
      } else {
        debug("group.skipped", "group skipped (parse returned null)", ctx);
      }
      break;
    }
    case "graphicFrame": {
      const gfResult = parseGraphicFrame(
        element,
        rels,
        slidePath,
        archive,
        colorResolver,
        fontScheme,
      );
      if (gfResult) {
        elements.push(gfResult);
      } else {
        debug("graphicFrame.skipped", "graphicFrame skipped (parse returned null)", ctx);
      }
      break;
    }
  }
}

function parseShape(
  sp: XmlNode,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
  orderedNode?: XmlOrderedNode,
  fmtScheme?: FormatScheme,
): ShapeElement | null {
  const spPr = sp.spPr as XmlNode | undefined;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm as XmlNode | undefined);
  if (!transform) return null;

  // Unsupported feature detection
  if (spPr.scene3d) {
    warn("spPr.scene3d", "3D scene/camera not implemented");
  }
  if (spPr.sp3d) {
    warn("spPr.sp3d", "3D extrusion/bevel not implemented");
  }

  const geometry = parseGeometry(spPr);

  // Resolve style references (sp.style) as fallback for direct attributes
  const styleRef = resolveShapeStyle(sp.style as XmlNode | undefined, fmtScheme, colorResolver);

  const directFill = parseFillFromNode(spPr, colorResolver, fillContext);
  const fill = directFill ?? styleRef?.fill ?? null;

  const directOutline = parseOutline(spPr.ln as XmlNode, colorResolver);
  const outline = directOutline ?? styleRef?.outline ?? null;

  // ordered ノードから txBody の ordered children を抽出
  let orderedTxBody: XmlOrderedNode[] | undefined;
  if (orderedNode) {
    const spChildren = orderedNode.sp as XmlOrderedNode[] | undefined;
    if (Array.isArray(spChildren)) {
      const txBodyEntry = spChildren.find((c: XmlOrderedNode) => "txBody" in c);
      orderedTxBody = txBodyEntry?.txBody as XmlOrderedNode[] | undefined;
    }
  }

  const textBody = parseTextBody(
    sp.txBody as XmlNode | undefined,
    colorResolver,
    rels,
    fontScheme,
    undefined,
    orderedTxBody,
  );

  const directEffects = parseEffectList(spPr.effectLst as XmlNode, colorResolver);
  const effects = directEffects ?? styleRef?.effects ?? null;

  const nvSpPr = sp.nvSpPr as XmlNode | undefined;
  const cNvPr = nvSpPr?.cNvPr as XmlNode | undefined;
  const altText = cNvPr?.["@_descr"] as string | undefined;
  const nvPr = nvSpPr?.nvPr as XmlNode | undefined;
  const ph = nvPr?.ph as XmlNode | undefined;
  const placeholderType = ph ? ((ph["@_type"] as string | undefined) ?? "body") : undefined;
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
    ...(altText && { altText }),
  };
}

function parseImage(
  pic: XmlNode,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): ImageElement | null {
  const spPr = pic.spPr as XmlNode | undefined;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm as XmlNode | undefined);
  if (!transform) return null;

  const blipFill = pic.blipFill as XmlNode | undefined;
  const blip = blipFill?.blip as XmlNode | undefined;
  const rId =
    (blip?.["@_r:embed"] as string | undefined) ?? (blip?.["@_embed"] as string | undefined);
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
  const imageData = uint8ArrayToBase64(mediaData);
  const effects = parseEffectList(spPr.effectLst as XmlNode, colorResolver);
  const blipEffects = parseBlipEffects(blip as XmlNode, colorResolver);
  const srcRect = parseSrcRect(blipFill?.srcRect as XmlNode | undefined);
  const stretch = parseStretchFillRect(blipFill?.stretch as XmlNode | undefined);
  const tile = parseTileInfo(blipFill?.tile as XmlNode | undefined);

  const nvPicPr = pic.nvPicPr as XmlNode | undefined;
  const cNvPr = nvPicPr?.cNvPr as XmlNode | undefined;
  const altText = cNvPr?.["@_descr"] as string | undefined;

  return {
    type: "image",
    transform,
    imageData,
    mimeType,
    effects,
    blipEffects,
    srcRect,
    ...(altText && { altText }),
    stretch,
    tile,
  };
}

function parseSrcRect(node: XmlNode | undefined): SrcRect | null {
  if (!node) return null;
  const l = Number(node["@_l"] ?? 0) / 100000;
  const t = Number(node["@_t"] ?? 0) / 100000;
  const r = Number(node["@_r"] ?? 0) / 100000;
  const b = Number(node["@_b"] ?? 0) / 100000;
  if (l === 0 && t === 0 && r === 0 && b === 0) return null;
  return { left: l, top: t, right: r, bottom: b };
}

function parseStretchFillRect(stretchNode: XmlNode | undefined): StretchFillRect | null {
  if (!stretchNode) return null;
  const fillRect = stretchNode.fillRect as XmlNode | undefined;
  if (!fillRect) return null;
  const l = Number(fillRect["@_l"] ?? 0) / 100000;
  const t = Number(fillRect["@_t"] ?? 0) / 100000;
  const r = Number(fillRect["@_r"] ?? 0) / 100000;
  const b = Number(fillRect["@_b"] ?? 0) / 100000;
  if (l === 0 && t === 0 && r === 0 && b === 0) return null;
  return { left: l, top: t, right: r, bottom: b };
}

function parseTileInfo(node: XmlNode | undefined): TileInfo | null {
  if (!node) return null;
  return {
    tx: Number(node["@_tx"] ?? 0),
    ty: Number(node["@_ty"] ?? 0),
    sx: Number(node["@_sx"] ?? 100000) / 100000,
    sy: Number(node["@_sy"] ?? 100000) / 100000,
    flip: ((node["@_flip"] as string) ?? "none") as TileInfo["flip"],
    align: (node["@_algn"] as string) ?? "tl",
  };
}

function parseConnector(
  cxn: XmlNode,
  colorResolver: ColorResolver,
  fmtScheme?: FormatScheme,
): ConnectorElement | null {
  const spPr = cxn.spPr as XmlNode | undefined;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm as XmlNode | undefined);
  if (!transform) return null;

  const geometry = parseGeometry(spPr);

  const styleRef = resolveShapeStyle(cxn.style as XmlNode | undefined, fmtScheme, colorResolver);
  const directOutline = parseOutline(spPr.ln as XmlNode, colorResolver);
  const outline = directOutline ?? styleRef?.outline ?? null;
  const directEffects = parseEffectList(spPr.effectLst as XmlNode, colorResolver);
  const effects = directEffects ?? styleRef?.effects ?? null;

  const nvCxnSpPr = cxn.nvCxnSpPr as XmlNode | undefined;
  const cNvPr = nvCxnSpPr?.cNvPr as XmlNode | undefined;
  const altText = cNvPr?.["@_descr"] as string | undefined;

  return { type: "connector", transform, geometry, outline, effects, ...(altText && { altText }) };
}

function parseGroup(
  grp: XmlNode,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  parentFillContext?: FillParseContext,
  fontScheme?: FontScheme | null,
  orderedChildren?: XmlOrderedNode[] | null,
  fmtScheme?: FormatScheme,
): GroupElement | null {
  const grpSpPr = grp.grpSpPr as XmlNode | undefined;
  if (!grpSpPr) return null;

  const xfrm = grpSpPr.xfrm as XmlNode | undefined;
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const childOff = xfrm?.chOff as XmlNode | undefined;
  const childExt = xfrm?.chExt as XmlNode | undefined;
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
    orderedChildren,
    fmtScheme,
  );
  const effects = parseEffectList(grpSpPr.effectLst as XmlNode, colorResolver);

  const nvGrpSpPr = grp.nvGrpSpPr as XmlNode | undefined;
  const cNvPr = nvGrpSpPr?.cNvPr as XmlNode | undefined;
  const altText = cNvPr?.["@_descr"] as string | undefined;

  return {
    type: "group",
    transform,
    childTransform,
    children,
    effects,
    ...(altText && { altText }),
  };
}

function parseGraphicFrame(
  gf: XmlNode,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): ChartElement | TableElement | GroupElement | null {
  const xfrm = gf.xfrm as XmlNode | undefined;
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const graphic = gf.graphic as XmlNode | undefined;
  const graphicData = graphic?.graphicData as XmlNode | undefined;
  if (!graphicData) return null;

  // Chart
  const chartRef = graphicData.chart as XmlNode | undefined;
  if (chartRef) {
    const rId =
      (chartRef["@_r:id"] as string | undefined) ?? (chartRef["@_id"] as string | undefined);
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
  const tblNode = graphicData.tbl as XmlNode | undefined;
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

  const uri = (graphicData["@_uri"] as string | undefined) ?? "unknown";
  warn("graphicFrame.unsupported", `unsupported graphicFrame content (uri: ${uri})`);

  return null;
}

function parseSmartArt(
  graphicData: XmlNode,
  transform: Transform,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): GroupElement | null {
  // dgm:relIds → removeNSPrefix → relIds
  const relIds = graphicData.relIds as XmlNode | undefined;
  if (!relIds) return null;

  // r:dm 属性 (data model relationship ID)
  const dmRId = (relIds["@_r:dm"] as string | undefined) ?? (relIds["@_dm"] as string | undefined);
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

  const parsed = parseXml(drawingXml);
  const drawing = parsed.drawing as XmlNode | undefined;
  const spTree = drawing?.spTree as XmlNode | undefined;
  if (!spTree) return null;

  // drawing XML 用のリレーションシップを取得
  const drawingRelsPath = buildRelsPath(drawingPath);
  const drawingRelsXml = archive.files.get(drawingRelsPath);
  const drawingRels = drawingRelsXml
    ? parseRelationships(drawingRelsXml)
    : new Map<string, Relationship>();

  // grpSpPr から childTransform を取得
  const grpSpPr = spTree.grpSpPr as XmlNode | undefined;
  const grpXfrm = grpSpPr?.xfrm as XmlNode | undefined;
  const childOff = grpXfrm?.chOff as XmlNode | undefined;
  const childExt = grpXfrm?.chExt as XmlNode | undefined;
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

  // ordered パーサーで子要素の出現順序を取得
  const orderedParsed = parseXmlOrdered(drawingXml);
  const orderedSpTree = navigateOrdered(orderedParsed, ["drawing", "spTree"]);

  const children = parseShapeTree(
    spTree,
    drawingRels,
    drawingPath,
    archive,
    colorResolver,
    undefined,
    fillContext,
    fontScheme,
    orderedSpTree,
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

function parseTransform(xfrm: XmlNode | undefined): Transform | null {
  if (!xfrm) return null;

  const off = xfrm.off as XmlNode | undefined;
  const ext = xfrm.ext as XmlNode | undefined;
  if (!off || !ext) return null;

  let offsetX = Number(off["@_x"] ?? 0);
  let offsetY = Number(off["@_y"] ?? 0);
  let extentWidth = Number(ext["@_cx"] ?? 0);
  let extentHeight = Number(ext["@_cy"] ?? 0);
  let rotation = Number(xfrm["@_rot"] ?? 0);

  if (Number.isNaN(offsetX)) {
    debug("transform.nan", "NaN detected in transform offsetX, defaulting to 0");
    offsetX = 0;
  }
  if (Number.isNaN(offsetY)) {
    debug("transform.nan", "NaN detected in transform offsetY, defaulting to 0");
    offsetY = 0;
  }
  if (Number.isNaN(extentWidth)) {
    debug("transform.nan", "NaN detected in transform extentWidth, defaulting to 0");
    extentWidth = 0;
  }
  if (Number.isNaN(extentHeight)) {
    debug("transform.nan", "NaN detected in transform extentHeight, defaulting to 0");
    extentHeight = 0;
  }
  if (Number.isNaN(rotation)) {
    debug("transform.nan", "NaN detected in transform rotation, defaulting to 0");
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

function parseGeometry(spPr: XmlNode): Geometry {
  if (spPr.prstGeom) {
    const prstGeom = spPr.prstGeom as XmlNode;
    const preset = (prstGeom["@_prst"] as string | undefined) ?? "rect";
    const avLst = prstGeom.avLst as XmlNode | undefined;
    const adjustValues: Record<string, number> = {};

    if (avLst?.gd) {
      const guides = Array.isArray(avLst.gd) ? avLst.gd : [avLst.gd];
      for (const gd of guides) {
        const gdNode = gd as XmlNode;
        const name = gdNode["@_name"] as string;
        const fmla = gdNode["@_fmla"] as string;
        const match = fmla?.match(/val\s+(\d+)/);
        if (name && match) {
          adjustValues[name] = Number(match[1]);
        }
      }
    }

    return { type: "preset", preset, adjustValues };
  }

  if (spPr.custGeom) {
    const paths = parseCustomGeometry(spPr.custGeom as XmlNode);
    if (paths) {
      return { type: "custom", paths };
    }
  }

  return { type: "preset", preset: "rect", adjustValues: {} };
}

export function parseTextBody(
  txBody: XmlNode | undefined,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
  lstStyleOverride?: DefaultTextStyle,
  orderedTxBody?: XmlOrderedNode[],
): TextBody | null {
  if (!txBody) return null;

  const bodyPr = txBody.bodyPr as XmlNode | undefined;

  const vert = (bodyPr?.["@_vert"] as BodyProperties["vert"] | undefined) ?? "horz";

  const numCol = bodyPr?.["@_numCol"] ? Math.max(1, Number(bodyPr["@_numCol"])) : 1;

  let autoFit: BodyProperties["autoFit"] = "noAutofit";
  let fontScale = 1;
  let lnSpcReduction = 0;
  if (bodyPr?.normAutofit !== undefined) {
    autoFit = "normAutofit";
    const normAutofit = bodyPr.normAutofit;
    if (typeof normAutofit === "object" && normAutofit !== null) {
      const normNode = normAutofit as XmlNode;
      fontScale = normNode["@_fontScale"] ? Number(normNode["@_fontScale"]) / 100000 : 1;
      lnSpcReduction = normNode["@_lnSpcReduction"]
        ? Number(normNode["@_lnSpcReduction"]) / 100000
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
    numCol,
    vert,
  };

  const lstStyle = lstStyleOverride ?? parseListStyle(txBody.lstStyle as XmlNode);

  // ordered data から各段落の ordered children を抽出
  const orderedParagraphs: (XmlOrderedNode[] | undefined)[] = [];
  if (orderedTxBody) {
    for (const child of orderedTxBody) {
      if ("p" in child) {
        orderedParagraphs.push(child.p as XmlOrderedNode[]);
      }
    }
  }

  const paragraphs: Paragraph[] = [];
  const pList = (txBody.p as XmlNode[] | undefined) ?? [];
  for (let i = 0; i < pList.length; i++) {
    paragraphs.push(
      parseParagraph(pList[i], colorResolver, rels, fontScheme, lstStyle, orderedParagraphs[i]),
    );
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

function parseBullet(pPr: XmlNode | undefined, colorResolver: ColorResolver) {
  let bullet: BulletType | null = null;
  let bulletFont: string | null = null;
  const buClr = pPr?.buClr as XmlNode | undefined;
  let bulletColor = colorResolver.resolve(buClr as XmlNode);
  if (!buClr) bulletColor = null;
  const buSzPct = pPr?.buSzPct as XmlNode | undefined;
  const bulletSizePct: number | null = buSzPct ? Number(buSzPct["@_val"]) : null;

  if (pPr?.buNone !== undefined) {
    bullet = { type: "none" };
  } else if (pPr?.buChar) {
    const buChar = pPr.buChar as XmlNode;
    bullet = { type: "char", char: (buChar["@_char"] as string | undefined) ?? "\u2022" };
  } else if (pPr?.buAutoNum) {
    const buAutoNum = pPr.buAutoNum as XmlNode;
    const scheme = (buAutoNum["@_type"] as string | undefined) ?? "arabicPeriod";
    bullet = {
      type: "autoNum",
      scheme: VALID_AUTO_NUM_SCHEMES.has(scheme) ? (scheme as AutoNumScheme) : "arabicPeriod",
      startAt: Number(buAutoNum["@_startAt"] ?? 1),
    };
  }

  if (pPr?.buFont) {
    const buFont = pPr.buFont as XmlNode;
    bulletFont = (buFont["@_typeface"] as string | undefined) ?? null;
  }

  return { bullet, bulletFont, bulletColor, bulletSizePct };
}

function parseSpacing(spc: XmlNode | undefined): SpacingValue {
  if (spc?.spcPts) {
    const spcPts = spc.spcPts as XmlNode;
    return { type: "pts", value: Number(spcPts["@_val"]) };
  }
  if (spc?.spcPct) {
    const spcPct = spc.spcPct as XmlNode;
    return { type: "pct", value: Number(spcPct["@_val"]) };
  }
  return { type: "pts", value: 0 };
}

function parseTabStops(pPr: XmlNode | undefined): TabStop[] {
  const tabLst = pPr?.tabLst as XmlNode | undefined;
  if (!tabLst) return [];

  const tabs = tabLst.tab;
  if (!tabs) return [];

  const tabArr = Array.isArray(tabs) ? (tabs as XmlNode[]) : [tabs as XmlNode];
  return tabArr.map((tab) => ({
    position: Number(tab["@_pos"] ?? 0),
    alignment: (tab["@_algn"] as TabStop["alignment"]) ?? "l",
  }));
}

/** 数式ノードからテキストを再帰的に抽出する */
function extractMathText(node: XmlNode): string {
  let text = "";
  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith("@_")) continue;
    if (key === "t") {
      if (typeof value === "object" && value !== null) {
        text += ((value as XmlNode)["#text"] as string) ?? "";
      } else if (value !== undefined && value !== null) {
        text += String(value as string | number);
      }
    } else if (typeof value === "object" && value !== null) {
      const items = Array.isArray(value) ? value : [value];
      for (const item of items) {
        if (typeof item === "object" && item !== null) {
          text += extractMathText(item as XmlNode);
        }
      }
    }
  }
  return text;
}

function extractTextContent(node: XmlNode): string {
  const text = node.t;
  if (typeof text === "object") {
    return ((text as XmlNode)["#text"] as string) ?? "";
  }
  return text !== undefined && text !== null ? String(text as string | number) : "";
}

function parseParagraph(
  p: XmlNode,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
  lstStyle?: DefaultTextStyle,
  orderedPChildren?: XmlOrderedNode[],
): Paragraph {
  const pPr = p.pPr as XmlNode | undefined;
  const level = Number(pPr?.["@_lvl"] ?? 0);

  // lstStyle からレベル対応のデフォルト段落プロパティを取得
  const lstLevelProps = lstStyle?.levels[level];

  const { bullet, bulletFont, bulletColor, bulletSizePct } = parseBullet(pPr, colorResolver);
  const lnSpc = pPr?.lnSpc as XmlNode | undefined;
  const lnSpcSpcPct = lnSpc?.spcPct as XmlNode | undefined;
  const tabStops = parseTabStops(pPr);
  const properties = {
    alignment: (pPr?.["@_algn"] as "l" | "ctr" | "r" | "just") ?? lstLevelProps?.alignment ?? "l",
    lineSpacing: lnSpcSpcPct ? Number(lnSpcSpcPct["@_val"]) : null,
    spaceBefore: parseSpacing(pPr?.spcBef as XmlNode | undefined),
    spaceAfter: parseSpacing(pPr?.spcAft as XmlNode | undefined),
    level,
    bullet,
    bulletFont,
    bulletColor,
    bulletSizePct,
    marginLeft:
      pPr?.["@_marL"] !== undefined ? Number(pPr["@_marL"]) : (lstLevelProps?.marginLeft ?? 0),
    indent:
      pPr?.["@_indent"] !== undefined ? Number(pPr["@_indent"]) : (lstLevelProps?.indent ?? 0),
    tabStops,
  };

  // defRPr のマージ: pPr.defRPr > lstStyle.lvl.defRPr
  const pPrDefRPr = parseDefaultRunProperties(pPr?.defRPr as XmlNode);
  const lstDefRPr = lstLevelProps?.defaultRunProperties;
  const mergedDefaults = mergeDefaultRunProperties(pPrDefRPr, lstDefRPr);

  const runs: TextRun[] = [];

  if (orderedPChildren) {
    // ordered children があれば、ドキュメント順で r/fld/br を処理
    const tagCounters: Record<string, number> = {};
    const rList = (p.r as XmlNode[] | undefined) ?? [];
    const fldList = (p.fld as XmlNode[] | undefined) ?? [];
    const brList = (p.br as XmlNode[] | undefined) ?? [];

    for (const child of orderedPChildren) {
      const tag = Object.keys(child).find((k) => k !== ":@");
      if (!tag) continue;

      const idx = tagCounters[tag] ?? 0;
      tagCounters[tag] = idx + 1;

      if (tag === "r") {
        const r = rList[idx];
        if (r) {
          const textContent = extractTextContent(r);
          const rPr = r.rPr as XmlNode | undefined;
          const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
          runs.push({ text: textContent, properties: runProps });
        }
      } else if (tag === "fld") {
        const fld = fldList[idx];
        if (fld) {
          const textContent = extractTextContent(fld);
          const rPr = fld.rPr as XmlNode | undefined;
          const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
          runs.push({ text: textContent, properties: runProps });
        }
      } else if (tag === "br") {
        const br = brList[idx];
        const rPr = br?.rPr as XmlNode | undefined;
        const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
        runs.push({ text: "\n", properties: runProps });
      } else if (tag === "m") {
        // 数式テキスト抽出 (a14:m → NS prefix removed → m)
        const mNode = p.m as XmlNode | XmlNode[] | undefined;
        if (mNode) {
          const mData = Array.isArray(mNode) ? mNode[idx] : idx === 0 ? mNode : undefined;
          if (mData) {
            const mathText = extractMathText(mData);
            if (mathText) {
              const runProps = parseRunProperties(
                undefined,
                colorResolver,
                rels,
                fontScheme,
                mergedDefaults,
              );
              runs.push({ text: mathText, properties: runProps });
            }
          }
        }
      }
    }
  } else {
    // フォールバック: ordered data がない場合は r のみ処理し、fld/br を追加
    const rList = (p.r as XmlNode[] | undefined) ?? [];
    for (const r of rList) {
      const textContent = extractTextContent(r);
      const rPr = r.rPr as XmlNode | undefined;
      const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
      runs.push({ text: textContent, properties: runProps });
    }

    // fld をフォールバック処理（ordered data なしでも最低限表示）
    const fldList = (p.fld as XmlNode[] | undefined) ?? [];
    for (const fld of fldList) {
      const textContent = extractTextContent(fld);
      const rPr = fld.rPr as XmlNode | undefined;
      const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
      runs.push({ text: textContent, properties: runProps });
    }

    // br をフォールバック処理
    const brList = (p.br as XmlNode[] | undefined) ?? [];
    for (const _br of brList) {
      const rPr = _br?.rPr as XmlNode | undefined;
      const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
      runs.push({ text: "\n", properties: runProps });
    }

    // 数式テキスト抽出
    const mNode = p.m as XmlNode | XmlNode[] | undefined;
    if (mNode) {
      const mArr = Array.isArray(mNode) ? mNode : [mNode];
      for (const m of mArr) {
        const mathText = extractMathText(m);
        if (mathText) {
          const runProps = parseRunProperties(
            undefined,
            colorResolver,
            rels,
            fontScheme,
            mergedDefaults,
          );
          runs.push({ text: mathText, properties: runProps });
        }
      }
    }
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

function parseRunProperties(
  rPr: XmlNode | undefined,
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

  // Unsupported feature detection
  if (rPr.effectLst) {
    warn("rPr.effectLst", "text run effects not implemented");
  }
  if (rPr.highlight) {
    warn("rPr.highlight", "text highlighting not implemented");
  }

  const latin = rPr.latin as XmlNode | undefined;
  const ea = rPr.ea as XmlNode | undefined;

  const fontSize = rPr["@_sz"]
    ? hundredthPointToPoint(Number(rPr["@_sz"]))
    : (defaults?.fontSize ?? null);
  const fontFamily = resolveThemeFont(
    (latin?.["@_typeface"] as string | undefined) ?? defaults?.fontFamily ?? null,
    fontScheme,
  );
  const fontFamilyEa = resolveThemeFont(
    (ea?.["@_typeface"] as string | undefined) ?? defaults?.fontFamilyEa ?? null,
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

  const solidFill = rPr.solidFill as XmlNode | undefined;
  let color = colorResolver.resolve(solidFill ?? rPr);
  if (!solidFill && !rPr.srgbClr && !rPr.schemeClr && !rPr.sysClr) {
    color = null;
  }

  const hyperlink = parseHyperlink(rPr.hlinkClick as XmlNode | undefined, rels);

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
  hlinkClick: XmlNode | undefined,
  rels?: Map<string, Relationship>,
): Hyperlink | null {
  if (!hlinkClick) return null;

  const rId =
    (hlinkClick["@_r:id"] as string | undefined) ?? (hlinkClick["@_id"] as string | undefined);
  if (!rId || !rels) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const tooltip = hlinkClick["@_tooltip"] as string | undefined;
  return { url: rel.target, ...(tooltip && { tooltip }) };
}
