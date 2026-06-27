import type { ChartElement } from "@pptx-glimpse/renderer";
import type { ImageElement, SrcRect, StretchFillRect, TileInfo } from "@pptx-glimpse/renderer";
import type {
  ConnectorElement,
  Geometry,
  GroupElement,
  ShapeElement,
  SlideElement,
  Transform,
} from "@pptx-glimpse/renderer";
import type { Background, Slide } from "@pptx-glimpse/renderer";
import type { TableElement } from "@pptx-glimpse/renderer";
import type {
  AutoNumScheme,
  BodyProperties,
  BulletType,
  DefaultRunProperties,
  DefaultTextStyle,
  Hyperlink,
  Paragraph,
  RunProperties,
  SpacingValue,
  TabStop,
  TextBody,
  TextOutline,
  TextRun,
} from "@pptx-glimpse/renderer";
import type { PlaceholderStyleInfo } from "@pptx-glimpse/renderer";
import type { FontScheme, FormatScheme } from "@pptx-glimpse/renderer";
import { uint8ArrayToBase64 } from "@pptx-glimpse/renderer";
import { hundredthPointToPoint } from "@pptx-glimpse/renderer";
import { asEmu, asHundredthPt } from "@pptx-glimpse/renderer";
import { debug, warn } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import { unsafeXmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { parseBlipEffects } from "./blip-effect-parser.js";
import { parseChart } from "./chart-parser.js";
import { parseCustomGeometry } from "./custom-geometry-parser.js";
import { parseEffectList } from "./effect-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import {
  buildRelsPath,
  parseRelationships,
  resolveRelationshipTarget,
} from "./relationship-parser.js";
import { resolveShapeStyle } from "./style-reference-resolver.js";
import { parseTable } from "./table-parser.js";
import {
  parseDefaultRunProperties,
  parseListStyle,
  resolveThemeFont,
} from "./text-style-parser.js";
import type { XmlNode, XmlOrderedNode } from "./xml-parser.js";
import { parseXml, parseXmlOrdered } from "./xml-parser.js";

const SHAPE_TAGS = new Set(["sp", "pic", "cxnSp", "grpSp", "graphicFrame"]);

const SMARTART_DIAGRAM_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "http://purl.oclc.org/ooxml/drawingml/diagram",
]);

// preserveOrder パース結果を path で辿り、指定ノードの子配列を返す。
export function navigateOrdered(
  ordered: XmlOrderedNode[],
  path: string[],
): XmlOrderedNode[] | null {
  let current: XmlOrderedNode[] = ordered;
  for (const key of path) {
    const entry = current.find((item: XmlOrderedNode) => key in item);
    if (!entry) return null;
    current = unsafeXmlBoundaryAssertion<XmlOrderedNode[]>(entry[key]);
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
  placeholderStyles?: PlaceholderStyleInfo[],
): Slide {
  const parsed = parseXml(slideXml);
  const sld = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.sld);
  if (!sld) {
    debug("slide.missing", `missing root element "sld" in XML`, `Slide ${slideNumber}`);
  }

  const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const fillContext: FillParseContext = { rels, archive, basePath: slidePath };
  const cSld = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sld?.cSld) ?? undefined;
  const background = parseBackground(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(cSld?.bg),
    colorResolver,
    fillContext,
  );

  // ordered パーサーで子要素の出現順序を取得（Z-order 保持のため）
  const orderedParsed = parseXmlOrdered(slideXml);
  const orderedSpTree = navigateOrdered(orderedParsed, ["sld", "cSld", "spTree"]);

  const elements = parseShapeTree(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(cSld?.spTree),
    rels,
    slidePath,
    archive,
    colorResolver,
    `Slide ${slideNumber}`,
    fillContext,
    fontScheme,
    orderedSpTree,
    fmtScheme,
    placeholderStyles,
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

  const bgPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(bgNode.bgPr);
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
      unsafeXmlBoundaryAssertion<unknown[]>(spTree[tag]).push(item);
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
  placeholderStyles?: PlaceholderStyleInfo[],
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
      placeholderStyles,
    );
  }

  // フォールバック: タイプ別イテレーション（後方互換）
  // mc:AlternateContent の処理: Choice 内の要素を spTree にマージ
  const alternateContents =
    unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.AlternateContent) ?? [];
  for (const ac of alternateContents) {
    const choices = Array.isArray(ac.Choice) ? ac.Choice : ac.Choice ? [ac.Choice] : [];
    for (const choice of choices) {
      mergeChildElements(spTree, unsafeXmlBoundaryAssertion<XmlNode>(choice));
    }
  }

  const ctx = context ?? slidePath;
  const elements: SlideElement[] = [];

  const shapes = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.sp) ?? [];
  for (const sp of shapes) {
    const shape = parseShape(
      sp,
      colorResolver,
      rels,
      fillContext,
      fontScheme,
      undefined,
      fmtScheme,
      placeholderStyles,
    );
    if (shape) {
      elements.push(shape);
    } else {
      debug("shape.skipped", "shape skipped (parse returned null)", ctx);
    }
  }

  const pics = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.pic) ?? [];
  for (const pic of pics) {
    const img = parseImage(pic, rels, slidePath, archive, colorResolver);
    if (img) {
      elements.push(img);
    } else {
      debug("image.skipped", "image skipped (parse returned null)", ctx);
    }
  }

  const cxnSps = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.cxnSp) ?? [];
  for (const cxn of cxnSps) {
    const connector = parseConnector(cxn, colorResolver, fmtScheme);
    if (connector) {
      elements.push(connector);
    } else {
      debug("connector.skipped", "connector skipped (parse returned null)", ctx);
    }
  }

  const grpSps = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.grpSp) ?? [];
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

  const graphicFrames =
    unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.graphicFrame) ?? [];
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
  placeholderStyles?: PlaceholderStyleInfo[],
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

      const acList =
        unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree.AlternateContent) ?? [];
      const acData = acList[acIndex] as XmlNode | undefined;
      if (!acData) continue;

      const choices = Array.isArray(acData.Choice)
        ? acData.Choice
        : acData.Choice
          ? [acData.Choice]
          : [];
      const firstChoice = unsafeXmlBoundaryAssertion<XmlNode | undefined>(choices[0]);
      if (!firstChoice) continue;

      // ordered 結果から Choice の子要素を取得
      const acOrderedChildren = unsafeXmlBoundaryAssertion<XmlOrderedNode[] | undefined>(
        child.AlternateContent,
      );

      const choiceOrdered = Array.isArray(acOrderedChildren)
        ? acOrderedChildren.find((c: XmlOrderedNode) => "Choice" in c)
        : null;
      const choiceChildren = unsafeXmlBoundaryAssertion<XmlOrderedNode[] | undefined>(
        choiceOrdered?.Choice,
      );

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
          const element = unsafeXmlBoundaryAssertion<XmlNode | undefined>(arr[choiceIdx]);
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
            placeholderStyles,
          );
        }
      }
      continue;
    }

    if (!SHAPE_TAGS.has(tag)) continue;

    const index = tagCounters[tag] ?? 0;
    tagCounters[tag] = index + 1;

    const tagArray = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(spTree[tag]);
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
      placeholderStyles,
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
  placeholderStyles?: PlaceholderStyleInfo[],
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
        placeholderStyles,
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
      const grpOrderedChildren =
        unsafeXmlBoundaryAssertion<XmlOrderedNode[] | undefined>(orderedNode[tag]) ?? null;
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
  placeholderStyles?: PlaceholderStyleInfo[],
): ShapeElement | null {
  const spPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.spPr);
  const spPrIsObject = spPr != null && typeof spPr === "object";

  // プレースホルダー情報を先に抽出（transform フォールバックに必要）
  const nvSpPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.nvSpPr);
  const nvPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvSpPr?.nvPr);
  const ph = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvPr?.ph);
  const placeholderType = ph
    ? (unsafeXmlBoundaryAssertion<string | undefined>(ph["@_type"]) ?? "body")
    : undefined;
  const placeholderIdx = ph?.["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;

  // spPr が空/未定義の場合、プレースホルダーならコンテキストからフォールバック
  let transform: Transform | null = null;
  let geometry: Geometry;

  if (spPrIsObject) {
    transform = parseTransform(unsafeXmlBoundaryAssertion<XmlNode | undefined>(spPr.xfrm));
    geometry = parseGeometry(spPr);
  } else {
    geometry = { type: "preset", preset: "rect", adjustValues: {} };
  }

  // transform が未解決の場合、プレースホルダーコンテキストからフォールバック
  if (!transform && placeholderType && placeholderStyles) {
    const inherited = findMatchingPlaceholder(placeholderType, placeholderIdx, placeholderStyles);
    if (inherited?.transform) {
      transform = inherited.transform;
    }
    if (!spPrIsObject && inherited?.geometry) {
      geometry = inherited.geometry;
    }
  }

  if (!transform) return null;

  // Unsupported feature detection
  if (spPrIsObject && spPr.scene3d) {
    warn("spPr.scene3d", "3D scene/camera not implemented");
  }
  if (spPrIsObject && spPr.sp3d) {
    warn("spPr.sp3d", "3D extrusion/bevel not implemented");
  }

  // Resolve style references (sp.style) as fallback for direct attributes
  const styleRef = resolveShapeStyle(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.style),
    fmtScheme,
    colorResolver,
  );

  const directFill = spPrIsObject ? parseFillFromNode(spPr, colorResolver, fillContext) : null;
  const fill = directFill ?? styleRef?.fill ?? null;

  const directOutline = spPrIsObject
    ? parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(spPr.ln), colorResolver)
    : null;
  const outline = directOutline ?? styleRef?.outline ?? null;

  // ordered ノードから txBody の ordered children を抽出
  let orderedTxBody: XmlOrderedNode[] | undefined;
  if (orderedNode) {
    const spChildren = unsafeXmlBoundaryAssertion<XmlOrderedNode[] | undefined>(orderedNode.sp);
    if (Array.isArray(spChildren)) {
      const txBodyEntry = spChildren.find((c: XmlOrderedNode) => "txBody" in c);
      orderedTxBody = unsafeXmlBoundaryAssertion<XmlOrderedNode[] | undefined>(txBodyEntry?.txBody);
    }
  }

  const textBody = parseTextBody(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(sp.txBody),
    colorResolver,
    rels,
    fontScheme,
    undefined,
    orderedTxBody,
  );

  const directEffects = spPrIsObject
    ? parseEffectList(unsafeXmlBoundaryAssertion<XmlNode>(spPr.effectLst), colorResolver)
    : null;
  const effects = directEffects ?? styleRef?.effects ?? null;

  const cNvPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvSpPr?.cNvPr);
  const altText = unsafeXmlBoundaryAssertion<string | undefined>(cNvPr?.["@_descr"]);
  const hyperlink = parseHyperlink(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(cNvPr?.hlinkClick),
    rels,
  );

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
    ...(hyperlink && { hyperlink }),
  };
}

function parseImage(
  pic: XmlNode,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): ImageElement | null {
  const spPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pic.spPr);
  if (!spPr) return null;

  const transform = parseTransform(unsafeXmlBoundaryAssertion<XmlNode | undefined>(spPr.xfrm));
  if (!transform) return null;

  const blipFill = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pic.blipFill);
  const blip = unsafeXmlBoundaryAssertion<XmlNode | undefined>(blipFill?.blip);
  const rId =
    unsafeXmlBoundaryAssertion<string | undefined>(blip?.["@_r:embed"]) ??
    unsafeXmlBoundaryAssertion<string | undefined>(blip?.["@_embed"]);
  if (!rId) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const mediaPath = resolveRelationshipTarget(slidePath, rel.target);
  const mediaData = archive.media.get(mediaPath);
  if (!mediaData) {
    debug("picture.media", `media file not found: ${mediaPath}`);
    return null;
  }

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
  const effects = parseEffectList(
    unsafeXmlBoundaryAssertion<XmlNode>(spPr.effectLst),
    colorResolver,
  );
  const blipEffects = parseBlipEffects(unsafeXmlBoundaryAssertion<XmlNode>(blip), colorResolver);
  const srcRect = parseSrcRect(unsafeXmlBoundaryAssertion<XmlNode | undefined>(blipFill?.srcRect));
  const stretch = parseStretchFillRect(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(blipFill?.stretch),
  );
  const tile = parseTileInfo(unsafeXmlBoundaryAssertion<XmlNode | undefined>(blipFill?.tile));

  const nvPicPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pic.nvPicPr);
  const cNvPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvPicPr?.cNvPr);
  const altText = unsafeXmlBoundaryAssertion<string | undefined>(cNvPr?.["@_descr"]);

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
  const fillRect = unsafeXmlBoundaryAssertion<XmlNode | undefined>(stretchNode.fillRect);
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
    tx: asEmu(Number(node["@_tx"] ?? 0)),
    ty: asEmu(Number(node["@_ty"] ?? 0)),
    sx: Number(node["@_sx"] ?? 100000) / 100000,
    sy: Number(node["@_sy"] ?? 100000) / 100000,
    flip: unsafeXmlBoundaryAssertion<TileInfo["flip"]>(node["@_flip"] ?? "none"),
    align: unsafeXmlBoundaryAssertion<string>(node["@_algn"]) ?? "tl",
  };
}

function parseConnector(
  cxn: XmlNode,
  colorResolver: ColorResolver,
  fmtScheme?: FormatScheme,
): ConnectorElement | null {
  const spPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cxn.spPr);
  if (!spPr) return null;

  const transform = parseTransform(unsafeXmlBoundaryAssertion<XmlNode | undefined>(spPr.xfrm));
  if (!transform) return null;

  const geometry = parseGeometry(spPr);

  const styleRef = resolveShapeStyle(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(cxn.style),
    fmtScheme,
    colorResolver,
  );
  const directOutline = parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(spPr.ln), colorResolver);
  const outline = directOutline ?? styleRef?.outline ?? null;
  const directEffects = parseEffectList(
    unsafeXmlBoundaryAssertion<XmlNode>(spPr.effectLst),
    colorResolver,
  );
  const effects = directEffects ?? styleRef?.effects ?? null;

  const nvCxnSpPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cxn.nvCxnSpPr);
  const cNvPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvCxnSpPr?.cNvPr);
  const altText = unsafeXmlBoundaryAssertion<string | undefined>(cNvPr?.["@_descr"]);

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
  const grpSpPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(grp.grpSpPr);
  if (!grpSpPr) return null;

  const xfrm = unsafeXmlBoundaryAssertion<XmlNode | undefined>(grpSpPr.xfrm);
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const childOff = unsafeXmlBoundaryAssertion<XmlNode | undefined>(xfrm?.chOff);
  const childExt = unsafeXmlBoundaryAssertion<XmlNode | undefined>(xfrm?.chExt);
  const childTransform: Transform = {
    offsetX: asEmu(Number(childOff?.["@_x"] ?? 0)),
    offsetY: asEmu(Number(childOff?.["@_y"] ?? 0)),
    extentWidth: asEmu(Number(childExt?.["@_cx"] ?? transform.extentWidth)),
    extentHeight: asEmu(Number(childExt?.["@_cy"] ?? transform.extentHeight)),
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
  const effects = parseEffectList(
    unsafeXmlBoundaryAssertion<XmlNode>(grpSpPr.effectLst),
    colorResolver,
  );

  const nvGrpSpPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(grp.nvGrpSpPr);
  const cNvPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(nvGrpSpPr?.cNvPr);
  const altText = unsafeXmlBoundaryAssertion<string | undefined>(cNvPr?.["@_descr"]);

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
  const xfrm = unsafeXmlBoundaryAssertion<XmlNode | undefined>(gf.xfrm);
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const graphic = unsafeXmlBoundaryAssertion<XmlNode | undefined>(gf.graphic);
  const graphicData = unsafeXmlBoundaryAssertion<XmlNode | undefined>(graphic?.graphicData);
  if (!graphicData) return null;

  // Chart
  const chartRef = unsafeXmlBoundaryAssertion<XmlNode | undefined>(graphicData.chart);
  if (chartRef) {
    const rId =
      unsafeXmlBoundaryAssertion<string | undefined>(chartRef["@_r:id"]) ??
      unsafeXmlBoundaryAssertion<string | undefined>(chartRef["@_id"]);
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
  const tblNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(graphicData.tbl);
  if (tblNode) {
    const tableData = parseTable(tblNode, colorResolver, fontScheme);
    if (!tableData) return null;

    return { type: "table", transform, table: tableData };
  }

  // SmartArt (Diagram) — Transitional and Strict URIs
  if (SMARTART_DIAGRAM_URIS.has(unsafeXmlBoundaryAssertion<string>(graphicData["@_uri"]))) {
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

  const uri = unsafeXmlBoundaryAssertion<string | undefined>(graphicData["@_uri"]) ?? "unknown";
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
  const relIds = unsafeXmlBoundaryAssertion<XmlNode | undefined>(graphicData.relIds);
  if (!relIds) return null;

  // r:dm 属性 (data model relationship ID)
  const dmRId =
    unsafeXmlBoundaryAssertion<string | undefined>(relIds["@_r:dm"]) ??
    unsafeXmlBoundaryAssertion<string | undefined>(relIds["@_dm"]);
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

  if (!drawingPath) {
    debug("smartArt.drawing", "diagramDrawing relationship not found");
    return null;
  }

  const drawingXml = archive.files.get(drawingPath);
  if (!drawingXml) {
    debug("smartArt.drawing", `drawing XML not found in archive: ${drawingPath}`);
    return null;
  }

  const parsed = parseXml(drawingXml);
  const drawing = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.drawing);
  const spTree = unsafeXmlBoundaryAssertion<XmlNode | undefined>(drawing?.spTree);
  if (!spTree) return null;

  // drawing XML 用のリレーションシップを取得
  const drawingRelsPath = buildRelsPath(drawingPath);
  const drawingRelsXml = archive.files.get(drawingRelsPath);
  const drawingRels = drawingRelsXml
    ? parseRelationships(drawingRelsXml)
    : new Map<string, Relationship>();

  // grpSpPr から childTransform を取得
  const grpSpPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(spTree.grpSpPr);
  const grpXfrm = unsafeXmlBoundaryAssertion<XmlNode | undefined>(grpSpPr?.xfrm);
  const childOff = unsafeXmlBoundaryAssertion<XmlNode | undefined>(grpXfrm?.chOff);
  const childExt = unsafeXmlBoundaryAssertion<XmlNode | undefined>(grpXfrm?.chExt);
  const childTransform: Transform = {
    offsetX: asEmu(Number(childOff?.["@_x"] ?? 0)),
    offsetY: asEmu(Number(childOff?.["@_y"] ?? 0)),
    extentWidth: asEmu(Number(childExt?.["@_cx"] ?? transform.extentWidth)),
    extentHeight: asEmu(Number(childExt?.["@_cy"] ?? transform.extentHeight)),
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

export function parseTransform(xfrm: XmlNode | undefined): Transform | null {
  if (!xfrm) return null;

  const off = unsafeXmlBoundaryAssertion<XmlNode | undefined>(xfrm.off);
  const ext = unsafeXmlBoundaryAssertion<XmlNode | undefined>(xfrm.ext);
  if (!off || !ext) return null;

  let offsetX = asEmu(Number(off["@_x"] ?? 0));
  let offsetY = asEmu(Number(off["@_y"] ?? 0));
  let extentWidth = asEmu(Number(ext["@_cx"] ?? 0));
  let extentHeight = asEmu(Number(ext["@_cy"] ?? 0));
  let rotation = Number(xfrm["@_rot"] ?? 0);

  if (Number.isNaN(offsetX)) {
    debug("transform.nan", "NaN detected in transform offsetX, defaulting to 0");
    offsetX = asEmu(0);
  }
  if (Number.isNaN(offsetY)) {
    debug("transform.nan", "NaN detected in transform offsetY, defaulting to 0");
    offsetY = asEmu(0);
  }
  if (Number.isNaN(extentWidth)) {
    debug("transform.nan", "NaN detected in transform extentWidth, defaulting to 0");
    extentWidth = asEmu(0);
  }
  if (Number.isNaN(extentHeight)) {
    debug("transform.nan", "NaN detected in transform extentHeight, defaulting to 0");
    extentHeight = asEmu(0);
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

export function parseGeometry(spPr: XmlNode): Geometry {
  if (spPr.prstGeom) {
    const prstGeom = unsafeXmlBoundaryAssertion<XmlNode>(spPr.prstGeom);
    const preset = unsafeXmlBoundaryAssertion<string | undefined>(prstGeom["@_prst"]) ?? "rect";
    const avLst = unsafeXmlBoundaryAssertion<XmlNode | undefined>(prstGeom.avLst);
    const adjustValues: Record<string, number> = {};

    if (avLst?.gd) {
      const guides = Array.isArray(avLst.gd) ? avLst.gd : [avLst.gd];
      for (const gd of guides) {
        const gdNode = unsafeXmlBoundaryAssertion<XmlNode>(gd);
        const name = unsafeXmlBoundaryAssertion<string>(gdNode["@_name"]);
        const fmla = unsafeXmlBoundaryAssertion<string>(gdNode["@_fmla"]);
        const match = fmla?.match(/val\s+(\d+)/);
        if (name && match) {
          adjustValues[name] = Number(match[1]);
        }
      }
    }

    return { type: "preset", preset, adjustValues };
  }

  if (spPr.custGeom) {
    const paths = parseCustomGeometry(unsafeXmlBoundaryAssertion<XmlNode>(spPr.custGeom));
    if (paths) {
      return { type: "custom", paths };
    }
  }

  return { type: "preset", preset: "rect", adjustValues: {} };
}

function findMatchingPlaceholder(
  placeholderType: string,
  placeholderIdx: number | undefined,
  styles: PlaceholderStyleInfo[],
): PlaceholderStyleInfo | undefined {
  // idx マッチ優先（transform を持つもの優先）
  if (placeholderIdx !== undefined) {
    const byIdx = styles.find((s) => s.placeholderIdx === placeholderIdx && s.transform);
    if (byIdx) return byIdx;
    // transform なしでも idx マッチがあれば返す
    const byIdxAny = styles.find((s) => s.placeholderIdx === placeholderIdx);
    if (byIdxAny) return byIdxAny;
  }

  // type マッチ（transform を持つもの優先）
  const byTypeWithTransform = styles.find(
    (s) => s.placeholderType === placeholderType && s.transform,
  );
  if (byTypeWithTransform) return byTypeWithTransform;

  // フォールバックタイプ (ctrTitle→title, subTitle→body) — transform 優先
  const fallbackType =
    placeholderType === "ctrTitle" ? "title" : placeholderType === "subTitle" ? "body" : undefined;
  if (fallbackType) {
    const byFallbackWithTransform = styles.find(
      (s) => s.placeholderType === fallbackType && s.transform,
    );
    if (byFallbackWithTransform) return byFallbackWithTransform;
  }

  // transform なしでも type マッチがあれば返す
  const byType = styles.find((s) => s.placeholderType === placeholderType);
  if (byType) return byType;

  if (fallbackType) {
    return styles.find((s) => s.placeholderType === fallbackType);
  }

  return undefined;
}

/**
 * orderedChildren 内に bullet を持つ pPr が2回以上出現するかを判定する。
 * 複数の pPr があっても bullet 開始が1回だけなら通常の parseParagraph で処理する。
 */
function hasMultipleBulletPPr(p: XmlNode, orderedChildren: XmlOrderedNode[]): boolean {
  const rawPPr = p.pPr;
  const pPrList: XmlNode[] = Array.isArray(rawPPr)
    ? unsafeXmlBoundaryAssertion<XmlNode[]>(rawPPr)
    : rawPPr
      ? [unsafeXmlBoundaryAssertion<XmlNode>(rawPPr)]
      : [];

  let bulletPPrCount = 0;
  let pPrCounter = 0;
  for (const child of orderedChildren) {
    const tag = Object.keys(child).find((k) => k !== ":@");
    if (tag === "pPr") {
      const pPrNode = pPrList[pPrCounter];
      if (
        pPrNode &&
        (pPrNode.buChar !== undefined ||
          pPrNode.buAutoNum !== undefined ||
          pPrNode.buBlip !== undefined)
      ) {
        bulletPPrCount++;
        if (bulletPPrCount >= 2) return true;
      }
      pPrCounter++;
    }
  }
  return false;
}

/**
 * 単一の <a:p> 内に複数の <a:pPr> が交互配置されている非標準XMLを
 * 複数の Paragraph に分割する。
 * buChar/buAutoNum/buBlip を持つ pPr を論理段落の開始として扱い、
 * buNone の pPr は同一段落内のスタイル変更として継続する。
 */
function splitInterleavedParagraph(
  p: XmlNode,
  orderedPChildren: XmlOrderedNode[],
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
  lstStyle?: DefaultTextStyle,
): Paragraph[] {
  const rawPPr = p.pPr;
  const pPrList: XmlNode[] = Array.isArray(rawPPr)
    ? unsafeXmlBoundaryAssertion<XmlNode[]>(rawPPr)
    : rawPPr
      ? [unsafeXmlBoundaryAssertion<XmlNode>(rawPPr)]
      : [];

  const rList: XmlNode[] = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.r) ?? [];
  const fldList: XmlNode[] = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.fld) ?? [];
  const brList: XmlNode[] = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.br) ?? [];

  interface Group {
    pPrIdx: number;
    rIndices: number[];
    fldIndices: number[];
    brIndices: number[];
    orderedChildren: XmlOrderedNode[];
  }

  const groups: Group[] = [];
  let currentGroup: Group | null = null;
  let pPrCounter = 0;
  const tagCounters: Record<string, number> = {};

  for (const child of orderedPChildren) {
    const tag = Object.keys(child).find((k) => k !== ":@");
    if (!tag) continue;

    if (tag === "pPr") {
      const pPrNode = pPrList[pPrCounter];
      const hasBullet =
        pPrNode &&
        (pPrNode.buChar !== undefined ||
          pPrNode.buAutoNum !== undefined ||
          pPrNode.buBlip !== undefined);

      if (hasBullet || !currentGroup) {
        if (currentGroup) groups.push(currentGroup);
        currentGroup = {
          pPrIdx: pPrCounter,
          rIndices: [],
          fldIndices: [],
          brIndices: [],
          orderedChildren: [child],
        };
      } else {
        currentGroup.orderedChildren.push(child);
      }
      pPrCounter++;
    } else if (tag === "endParaRPr") {
      // endParaRPr は最後のグループに付与（後で処理）
    } else {
      const globalIdx = tagCounters[tag] ?? 0;
      tagCounters[tag] = globalIdx + 1;

      if (!currentGroup) {
        currentGroup = {
          pPrIdx: -1,
          rIndices: [],
          fldIndices: [],
          brIndices: [],
          orderedChildren: [],
        };
      }

      if (tag === "r") currentGroup.rIndices.push(globalIdx);
      else if (tag === "fld") currentGroup.fldIndices.push(globalIdx);
      else if (tag === "br") currentGroup.brIndices.push(globalIdx);
      currentGroup.orderedChildren.push(child);
    }
  }
  if (currentGroup) groups.push(currentGroup);

  const results: Paragraph[] = [];
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    const isLast = g === groups.length - 1;

    const syntheticP: XmlNode = {};
    if (group.pPrIdx >= 0 && pPrList[group.pPrIdx]) {
      syntheticP.pPr = pPrList[group.pPrIdx];
    }
    if (group.rIndices.length > 0) {
      syntheticP.r = group.rIndices.map((i) => rList[i]);
    }
    if (group.fldIndices.length > 0) {
      syntheticP.fld = group.fldIndices.map((i) => fldList[i]);
    }
    if (group.brIndices.length > 0) {
      syntheticP.br = group.brIndices.map((i) => brList[i]);
    }
    if (isLast && p.endParaRPr) {
      syntheticP.endParaRPr = p.endParaRPr;
    }

    const paragraph = parseParagraph(
      syntheticP,
      colorResolver,
      rels,
      fontScheme,
      lstStyle,
      group.orderedChildren,
    );

    // 非最終グループの末尾改行をストリップ（非標準フォーマットで段落区切りとして使われている）
    // 最終グループは改行を保持する（意図的な改行の可能性がある）
    if (!isLast && paragraph.runs.length > 0) {
      const lastRun = paragraph.runs[paragraph.runs.length - 1];
      if (lastRun.text.endsWith("\n")) {
        paragraph.runs[paragraph.runs.length - 1] = {
          ...lastRun,
          text: lastRun.text.slice(0, -1),
        };
      }
    }

    results.push(paragraph);
  }

  return results;
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

  const bodyPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(txBody.bodyPr);

  const vert =
    unsafeXmlBoundaryAssertion<BodyProperties["vert"] | undefined>(bodyPr?.["@_vert"]) ?? "horz";

  const numCol = bodyPr?.["@_numCol"] ? Math.max(1, Number(bodyPr["@_numCol"])) : 1;

  let autoFit: BodyProperties["autoFit"] = "noAutofit";
  let fontScale = 1;
  let lnSpcReduction = 0;
  if (bodyPr?.normAutofit !== undefined) {
    autoFit = "normAutofit";
    const normAutofit = bodyPr.normAutofit;
    if (typeof normAutofit === "object" && normAutofit !== null) {
      const normNode = unsafeXmlBoundaryAssertion<XmlNode>(normAutofit);
      fontScale = normNode["@_fontScale"] ? Number(normNode["@_fontScale"]) / 100000 : 1;
      lnSpcReduction = normNode["@_lnSpcReduction"]
        ? Number(normNode["@_lnSpcReduction"]) / 100000
        : 0;
    }
  } else if (bodyPr?.spAutoFit !== undefined) {
    autoFit = "spAutofit";
  }

  const bodyProperties: BodyProperties = {
    anchor: unsafeXmlBoundaryAssertion<"t" | "ctr" | "b">(bodyPr?.["@_anchor"]) ?? "t",
    marginLeft: asEmu(Number(bodyPr?.["@_lIns"] ?? 91440)),
    marginRight: asEmu(Number(bodyPr?.["@_rIns"] ?? 91440)),
    marginTop: asEmu(Number(bodyPr?.["@_tIns"] ?? 45720)),
    marginBottom: asEmu(Number(bodyPr?.["@_bIns"] ?? 45720)),
    wrap: unsafeXmlBoundaryAssertion<"square" | "none">(bodyPr?.["@_wrap"]) ?? "square",
    autoFit,
    fontScale,
    lnSpcReduction,
    numCol,
    vert,
  };

  const lstStyle =
    lstStyleOverride ?? parseListStyle(unsafeXmlBoundaryAssertion<XmlNode>(txBody.lstStyle));

  // ordered data から各段落の ordered children を抽出
  const orderedParagraphs: (XmlOrderedNode[] | undefined)[] = [];
  if (orderedTxBody) {
    for (const child of orderedTxBody) {
      if ("p" in child) {
        orderedParagraphs.push(unsafeXmlBoundaryAssertion<XmlOrderedNode[]>(child.p));
      }
    }
  }

  const paragraphs: Paragraph[] = [];
  const pList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(txBody.p) ?? [];
  for (let i = 0; i < pList.length; i++) {
    const orderedChildren = orderedParagraphs[i];
    // 単一 <a:p> 内に bullet pPr が複数回交互配置されている非標準XMLを検出して分割
    // bullet を持つ pPr が2回以上出現するケースのみ分割対象にする
    if (orderedChildren && hasMultipleBulletPPr(pList[i], orderedChildren)) {
      paragraphs.push(
        ...splitInterleavedParagraph(
          pList[i],
          orderedChildren,
          colorResolver,
          rels,
          fontScheme,
          lstStyle,
        ),
      );
    } else {
      paragraphs.push(
        parseParagraph(pList[i], colorResolver, rels, fontScheme, lstStyle, orderedChildren),
      );
    }
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
  const buClr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pPr?.buClr);
  let bulletColor = colorResolver.resolve(unsafeXmlBoundaryAssertion<XmlNode>(buClr));
  if (!buClr) bulletColor = null;
  const buSzPct = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pPr?.buSzPct);
  const bulletSizePct: number | null = buSzPct ? Number(buSzPct["@_val"]) : null;

  if (pPr?.buNone !== undefined) {
    bullet = { type: "none" };
  } else if (pPr?.buChar) {
    const buChar = unsafeXmlBoundaryAssertion<XmlNode>(pPr.buChar);
    bullet = {
      type: "char",
      char: unsafeXmlBoundaryAssertion<string | undefined>(buChar["@_char"]) ?? "\u2022",
    };
  } else if (pPr?.buAutoNum) {
    const buAutoNum = unsafeXmlBoundaryAssertion<XmlNode>(pPr.buAutoNum);
    const scheme =
      unsafeXmlBoundaryAssertion<string | undefined>(buAutoNum["@_type"]) ?? "arabicPeriod";
    bullet = {
      type: "autoNum",
      scheme: VALID_AUTO_NUM_SCHEMES.has(scheme)
        ? unsafeXmlBoundaryAssertion<AutoNumScheme>(scheme)
        : "arabicPeriod",
      startAt: Number(buAutoNum["@_startAt"] ?? 1),
    };
  }

  if (pPr?.buFont) {
    const buFont = unsafeXmlBoundaryAssertion<XmlNode>(pPr.buFont);
    bulletFont = unsafeXmlBoundaryAssertion<string | undefined>(buFont["@_typeface"]) ?? null;
  }

  return { bullet, bulletFont, bulletColor, bulletSizePct };
}

function parseSpacing(spc: XmlNode | undefined): SpacingValue {
  if (spc?.spcPts) {
    const spcPts = unsafeXmlBoundaryAssertion<XmlNode>(spc.spcPts);
    return { type: "pts", value: asHundredthPt(Number(spcPts["@_val"])) };
  }
  if (spc?.spcPct) {
    const spcPct = unsafeXmlBoundaryAssertion<XmlNode>(spc.spcPct);
    return { type: "pct", value: Number(spcPct["@_val"]) };
  }
  return { type: "pts", value: asHundredthPt(0) };
}

function parseTabStops(pPr: XmlNode | undefined): TabStop[] {
  const tabLst = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pPr?.tabLst);
  if (!tabLst) return [];

  const tabs = tabLst.tab;
  if (!tabs) return [];

  const tabArr = Array.isArray(tabs)
    ? unsafeXmlBoundaryAssertion<XmlNode[]>(tabs)
    : [unsafeXmlBoundaryAssertion<XmlNode>(tabs)];
  return tabArr.map((tab) => ({
    position: asEmu(Number(tab["@_pos"] ?? 0)),
    alignment: unsafeXmlBoundaryAssertion<TabStop["alignment"]>(tab["@_algn"]) ?? "l",
  }));
}

/** 数式ノードからテキストを再帰的に抽出する */
function extractMathText(node: XmlNode): string {
  let text = "";
  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith("@_")) continue;
    if (key === "t") {
      if (typeof value === "object" && value !== null) {
        text +=
          unsafeXmlBoundaryAssertion<string>(unsafeXmlBoundaryAssertion<XmlNode>(value)["#text"]) ??
          "";
      } else if (
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean"
      ) {
        text += String(value);
      }
    } else if (typeof value === "object" && value !== null) {
      const items = Array.isArray(value) ? value : [value];
      for (const item of items) {
        if (typeof item === "object" && item !== null) {
          text += extractMathText(unsafeXmlBoundaryAssertion<XmlNode>(item));
        }
      }
    }
  }
  return text;
}

function extractTextContent(node: XmlNode): string {
  const text = node.t;
  if (typeof text === "object") {
    return (
      unsafeXmlBoundaryAssertion<string>(unsafeXmlBoundaryAssertion<XmlNode>(text)["#text"]) ?? ""
    );
  }
  return typeof text === "string" || typeof text === "number" || typeof text === "boolean"
    ? String(text)
    : "";
}

function parseParagraph(
  p: XmlNode,
  colorResolver: ColorResolver,
  rels?: Map<string, Relationship>,
  fontScheme?: FontScheme | null,
  lstStyle?: DefaultTextStyle,
  orderedPChildren?: XmlOrderedNode[],
): Paragraph {
  const pPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(p.pPr);
  const level = Number(pPr?.["@_lvl"] ?? 0);

  // lstStyle からレベル対応のデフォルト段落プロパティを取得
  const lstLevelProps = lstStyle?.levels[level];

  const { bullet, bulletFont, bulletColor, bulletSizePct } = parseBullet(pPr, colorResolver);
  const lnSpc = unsafeXmlBoundaryAssertion<XmlNode | undefined>(pPr?.lnSpc);
  const tabStops = parseTabStops(pPr);
  const properties = {
    alignment:
      unsafeXmlBoundaryAssertion<"l" | "ctr" | "r" | "just">(pPr?.["@_algn"]) ??
      lstLevelProps?.alignment ??
      null,
    lineSpacing: lnSpc?.spcPts || lnSpc?.spcPct ? parseSpacing(lnSpc) : null,
    spaceBefore: parseSpacing(unsafeXmlBoundaryAssertion<XmlNode | undefined>(pPr?.spcBef)),
    spaceAfter: parseSpacing(unsafeXmlBoundaryAssertion<XmlNode | undefined>(pPr?.spcAft)),
    level,
    bullet: bullet ?? lstLevelProps?.bullet ?? null,
    bulletFont: bulletFont ?? lstLevelProps?.bulletFont ?? null,
    bulletColor: bulletColor ?? lstLevelProps?.bulletColor ?? null,
    bulletSizePct: bulletSizePct ?? lstLevelProps?.bulletSizePct ?? null,
    marginLeft:
      pPr?.["@_marL"] !== undefined
        ? asEmu(Number(pPr["@_marL"]))
        : (lstLevelProps?.marginLeft ?? null),
    indent:
      pPr?.["@_indent"] !== undefined
        ? asEmu(Number(pPr["@_indent"]))
        : (lstLevelProps?.indent ?? null),
    tabStops,
  };

  // defRPr のマージ: pPr.defRPr > lstStyle.lvl.defRPr
  const pPrDefRPr = parseDefaultRunProperties(unsafeXmlBoundaryAssertion<XmlNode>(pPr?.defRPr));
  const lstDefRPr = lstLevelProps?.defaultRunProperties;
  const mergedDefaults = mergeDefaultRunProperties(pPrDefRPr, lstDefRPr);

  const runs: TextRun[] = [];

  if (orderedPChildren) {
    // ordered children があれば、ドキュメント順で r/fld/br を処理
    const tagCounters: Record<string, number> = {};
    const rList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.r) ?? [];
    const fldList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.fld) ?? [];
    const brList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.br) ?? [];

    for (const child of orderedPChildren) {
      const tag = Object.keys(child).find((k) => k !== ":@");
      if (!tag) continue;

      const idx = tagCounters[tag] ?? 0;
      tagCounters[tag] = idx + 1;

      if (tag === "r") {
        const r = rList[idx];
        if (r) {
          const textContent = extractTextContent(r);
          const rPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(r.rPr);
          const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
          runs.push({ text: textContent, properties: runProps });
        }
      } else if (tag === "fld") {
        const fld = fldList[idx];
        if (fld) {
          const textContent = extractTextContent(fld);
          const rPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(fld.rPr);
          const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
          runs.push({ text: textContent, properties: runProps });
        }
      } else if (tag === "br") {
        const br = brList[idx];
        const rPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(br?.rPr);
        const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
        runs.push({ text: "\n", properties: runProps });
      } else if (tag === "m") {
        // 数式テキスト抽出 (a14:m → NS prefix removed → m)
        const mNode = unsafeXmlBoundaryAssertion<XmlNode | XmlNode[] | undefined>(p.m);
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
    const rList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.r) ?? [];
    for (const r of rList) {
      const textContent = extractTextContent(r);
      const rPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(r.rPr);
      const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
      runs.push({ text: textContent, properties: runProps });
    }

    // fld をフォールバック処理（ordered data なしでも最低限表示）
    const fldList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.fld) ?? [];
    for (const fld of fldList) {
      const textContent = extractTextContent(fld);
      const rPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(fld.rPr);
      const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
      runs.push({ text: textContent, properties: runProps });
    }

    // br をフォールバック処理
    const brList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.br) ?? [];
    for (const _br of brList) {
      const rPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(_br?.rPr);
      const runProps = parseRunProperties(rPr, colorResolver, rels, fontScheme, mergedDefaults);
      runs.push({ text: "\n", properties: runProps });
    }

    // 数式テキスト抽出
    const mNode = unsafeXmlBoundaryAssertion<XmlNode | XmlNode[] | undefined>(p.m);
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

  // endParaRPr をパース（空段落の高さ計算に使用）
  const endParaRPrNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(p.endParaRPr);
  const endParaRunProperties = endParaRPrNode
    ? parseRunProperties(endParaRPrNode, colorResolver, rels, fontScheme, mergedDefaults)
    : undefined;

  return { runs, properties, endParaRunProperties };
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
    fontFamilyCs: primary.fontFamilyCs ?? secondary.fontFamilyCs,
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
      fontFamilyCs: resolveThemeFont(defaults?.fontFamilyCs ?? null, fontScheme),
      bold: defaults?.bold ?? false,
      italic: defaults?.italic ?? false,
      underline: defaults?.underline ?? false,
      strikethrough: defaults?.strikethrough ?? false,
      color: null,
      baseline: 0,
      hyperlink: null,
      outline: null,
    };
  }

  // Unsupported feature detection
  if (rPr.effectLst) {
    warn("rPr.effectLst", "text run effects not implemented");
  }
  if (rPr.highlight) {
    warn("rPr.highlight", "text highlighting not implemented");
  }

  const latin = unsafeXmlBoundaryAssertion<XmlNode | undefined>(rPr.latin);
  const ea = unsafeXmlBoundaryAssertion<XmlNode | undefined>(rPr.ea);
  const cs = unsafeXmlBoundaryAssertion<XmlNode | undefined>(rPr.cs);

  const fontSize = rPr["@_sz"]
    ? hundredthPointToPoint(asHundredthPt(Number(rPr["@_sz"])))
    : (defaults?.fontSize ?? null);
  const fontFamily = resolveThemeFont(
    unsafeXmlBoundaryAssertion<string | undefined>(latin?.["@_typeface"]) ??
      defaults?.fontFamily ??
      null,
    fontScheme,
  );
  const fontFamilyEa = resolveThemeFont(
    unsafeXmlBoundaryAssertion<string | undefined>(ea?.["@_typeface"]) ??
      defaults?.fontFamilyEa ??
      null,
    fontScheme,
  );
  const fontFamilyCs = resolveThemeFont(
    unsafeXmlBoundaryAssertion<string | undefined>(cs?.["@_typeface"]) ??
      defaults?.fontFamilyCs ??
      null,
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
  const hasExplicitUnderline = rPr["@_u"] !== undefined;
  let underline = hasExplicitUnderline ? rPr["@_u"] !== "none" : (defaults?.underline ?? false);
  const strikethrough =
    rPr["@_strike"] !== undefined
      ? rPr["@_strike"] !== "noStrike"
      : (defaults?.strikethrough ?? false);
  const baseline = rPr["@_baseline"] ? Number(rPr["@_baseline"]) / 1000 : 0;

  const solidFill = unsafeXmlBoundaryAssertion<XmlNode | undefined>(rPr.solidFill);
  let color = colorResolver.resolve(solidFill ?? rPr);
  if (!solidFill && !rPr.srgbClr && !rPr.schemeClr && !rPr.sysClr) {
    color = null;
  }

  const hyperlink = parseHyperlink(
    unsafeXmlBoundaryAssertion<XmlNode | undefined>(rPr.hlinkClick),
    rels,
  );

  // Default hyperlink style: theme hlink color + underline
  if (hyperlink) {
    if (!color) {
      color = colorResolver.resolve({ schemeClr: { "@_val": "hlink" } });
    }
    if (!hasExplicitUnderline && !underline) {
      underline = true;
    }
  }

  const ln = unsafeXmlBoundaryAssertion<XmlNode | undefined>(rPr.ln);
  let outline: TextOutline | null = null;
  if (ln) {
    const lnWidth = asEmu(Number(ln["@_w"] ?? 12700));
    const lnFill = unsafeXmlBoundaryAssertion<XmlNode | undefined>(ln.solidFill);
    const lnColor = lnFill ? colorResolver.resolve(lnFill) : null;
    if (lnColor) {
      outline = { width: lnWidth, color: lnColor };
    }
  }

  return {
    fontSize,
    fontFamily,
    fontFamilyEa,
    fontFamilyCs,
    bold,
    italic,
    underline,
    strikethrough,
    color,
    baseline,
    hyperlink,
    outline,
  };
}

function parseHyperlink(
  hlinkClick: XmlNode | undefined,
  rels?: Map<string, Relationship>,
): Hyperlink | null {
  if (!hlinkClick) return null;

  const rId =
    unsafeXmlBoundaryAssertion<string | undefined>(hlinkClick["@_r:id"]) ??
    unsafeXmlBoundaryAssertion<string | undefined>(hlinkClick["@_id"]);
  if (!rId || !rels) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const tooltip = unsafeXmlBoundaryAssertion<string | undefined>(hlinkClick["@_tooltip"]);
  return { url: rel.target, ...(tooltip && { tooltip }) };
}
