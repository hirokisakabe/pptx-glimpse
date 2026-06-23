/**
 * `p:txBody` を CleanDoc source の text body / paragraph / run へ読み取る。
 *
 * PoC scope は plain run text と basic run properties (太字 / 斜体 / 下線 /
 * フォントサイズ / typeface / solid color / basic bullet) を typed に表す
 * (`docs/cleandoc-minimal-poc-scope.md`)。field / line break 等の未対応ノードは
 * raw sidecar として保持する。
 */

import type {
  PartPath,
  RawSidecarId,
  SourceNodeId,
  SourceParagraph,
  SourceParagraphProperties,
  SourceRunProperties,
  SourceTextAlign,
  SourceTextBody,
  SourceTextBodyProperties,
  SourceTextRun,
  SourceVerticalAnchor,
} from "../source/index.js";
import { asEmu, asHundredthPt, asPt, asSourceNodeId } from "../source/index.js";
import { parseColorElement } from "./drawing.js";
import { isTrue, numericAttr } from "./drawing.js";
import { collectUnknownSidecars } from "./raw-node.js";
import { getAttr, getChild, getChildArray, getChildText, type XmlNode } from "./xml.js";

const KNOWN_TXBODY_CHILDREN: ReadonlySet<string> = new Set(["bodyPr", "lstStyle", "p"]);
const KNOWN_PARAGRAPH_CHILDREN: ReadonlySet<string> = new Set(["pPr", "r"]);
const KNOWN_RUN_CHILDREN: ReadonlySet<string> = new Set(["rPr", "t"]);
const KNOWN_RUN_PROPERTY_CHILDREN: ReadonlySet<string> = new Set(["latin", "solidFill"]);

const ALIGN_MAP: Readonly<Record<string, SourceTextAlign>> = {
  l: "left",
  ctr: "center",
  r: "right",
  just: "justify",
};

const ANCHOR_MAP: Readonly<Record<string, SourceVerticalAnchor>> = {
  t: "top",
  ctr: "middle",
  b: "bottom",
};

/** `p:txBody` を読む。`txBody` が無い shape では undefined。 */
export function parseTextBody(
  txBody: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
): SourceTextBody | undefined {
  if (!txBody) return undefined;

  const properties = parseBodyProperties(getChild(txBody, "bodyPr"));
  const paragraphs = getChildArray(txBody, "p").map((p, paragraphIndex) =>
    parseParagraph(p, partPath, nextId, ownerNodeId, ownerOrderingSlot, paragraphIndex),
  );
  const rawSidecars = collectUnknownSidecars(txBody, KNOWN_TXBODY_CHILDREN, nextId);

  return {
    paragraphs,
    ...(properties !== undefined ? { properties } : {}),
    handle: { partPath },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseBodyProperties(bodyPr: XmlNode | undefined): SourceTextBodyProperties | undefined {
  if (!bodyPr) return undefined;

  const marginLeft = numericAttr(bodyPr, "lIns");
  const marginRight = numericAttr(bodyPr, "rIns");
  const marginTop = numericAttr(bodyPr, "tIns");
  const marginBottom = numericAttr(bodyPr, "bIns");
  const anchorToken = getAttr(bodyPr, "anchor");
  const anchor = anchorToken !== undefined ? ANCHOR_MAP[anchorToken] : undefined;

  const properties: SourceTextBodyProperties = {
    ...(marginLeft !== undefined ? { marginLeft: emu(marginLeft) } : {}),
    ...(marginRight !== undefined ? { marginRight: emu(marginRight) } : {}),
    ...(marginTop !== undefined ? { marginTop: emu(marginTop) } : {}),
    ...(marginBottom !== undefined ? { marginBottom: emu(marginBottom) } : {}),
    ...(anchor !== undefined ? { anchor } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

function parseParagraph(
  p: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
): SourceParagraph {
  const properties = parseParagraphProperties(getChild(p, "pPr"));
  const runs = getChildArray(p, "r").map((r, runIndex) =>
    parseRun(r, partPath, nextId, ownerNodeId, ownerOrderingSlot, paragraphIndex, runIndex),
  );
  const rawSidecars = collectUnknownSidecars(p, KNOWN_PARAGRAPH_CHILDREN, nextId);

  return {
    runs,
    ...(properties !== undefined ? { properties } : {}),
    handle: {
      partPath,
      nodeId: textNodeId("paragraph", ownerNodeId, ownerOrderingSlot, paragraphIndex),
      orderingSlot: paragraphIndex,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseParagraphProperties(pPr: XmlNode | undefined): SourceParagraphProperties | undefined {
  if (!pPr) return undefined;

  const alignToken = getAttr(pPr, "algn");
  const align = alignToken !== undefined ? ALIGN_MAP[alignToken] : undefined;
  const level = numericAttr(pPr, "lvl");
  const lineSpacingPts = numericAttr(getChild(getChild(pPr, "lnSpc"), "spcPts"), "val");
  const bullet = parseBullet(pPr);
  const bulletFont = getAttr(getChild(pPr, "buFont"), "typeface");

  const properties: SourceParagraphProperties = {
    ...(align !== undefined ? { align } : {}),
    ...(level !== undefined ? { level } : {}),
    ...(lineSpacingPts !== undefined ? { lineSpacingPts: asHundredthPt(lineSpacingPts) } : {}),
    ...(bullet !== undefined ? { bullet } : {}),
    ...(bulletFont !== undefined ? { bulletFont } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

function parseBullet(pPr: XmlNode): SourceParagraphProperties["bullet"] {
  if (getChild(pPr, "buNone") !== undefined) return { type: "none" };

  // Numbered bullets (`a:buAutoNum`) stay outside the current basic bullet subset.
  const char = getAttr(getChild(pPr, "buChar"), "char");
  if (char) return { type: "char", char };

  return undefined;
}

function parseRun(
  r: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  runIndex: number,
): SourceTextRun {
  const properties = parseRunProperties(getChild(r, "rPr"));
  const rawSidecars = collectRunSidecars(r, nextId);

  return {
    kind: "textRun",
    text: getChildText(r, "t") ?? "",
    ...(properties !== undefined ? { properties } : {}),
    handle: {
      partPath,
      nodeId: textNodeId("run", ownerNodeId, ownerOrderingSlot, paragraphIndex, runIndex),
      orderingSlot: runIndex,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseRunProperties(rPr: XmlNode | undefined): SourceRunProperties | undefined {
  if (!rPr) return undefined;

  const bold = getAttr(rPr, "b");
  const italic = getAttr(rPr, "i");
  const underline = getAttr(rPr, "u");
  const size = numericAttr(rPr, "sz");
  const typeface = getAttr(getChild(rPr, "latin"), "typeface");
  const color = parseColorElement(getChild(rPr, "solidFill"));

  const properties: SourceRunProperties = {
    ...(bold !== undefined ? { bold: isTrue(bold) } : {}),
    ...(italic !== undefined ? { italic: isTrue(italic) } : {}),
    ...(underline !== undefined ? { underline: underline !== "none" } : {}),
    // `a:rPr@sz` は 1/100 pt 単位。pt へ変換して保持する。
    ...(size !== undefined ? { fontSize: asPt(size / 100) } : {}),
    ...(typeface !== undefined ? { typeface } : {}),
    ...(color !== undefined ? { color } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

/**
 * run の未対応素材を sidecar に集める。run 直下 (`a:rPr` / `a:t` 以外) に加え、
 * `a:rPr` 内の未対応子要素 (`a:ea` / `a:cs` / `a:hlinkClick` 等) も保持する。
 */
function collectRunSidecars(
  r: XmlNode,
  nextId: () => RawSidecarId,
): ReturnType<typeof collectUnknownSidecars> {
  const runLevel = collectUnknownSidecars(r, KNOWN_RUN_CHILDREN, nextId);
  const propLevel = collectUnknownSidecars(getChild(r, "rPr"), KNOWN_RUN_PROPERTY_CHILDREN, nextId);
  return [...runLevel, ...propLevel];
}

function emu(value: number) {
  return asEmu(value);
}

function textNodeId(
  kind: "paragraph" | "run",
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  runIndex?: number,
): SourceNodeId {
  const owner =
    ownerNodeId !== undefined ? `shape:${ownerNodeId}` : `shapeSlot:${ownerOrderingSlot}`;
  const suffix = kind === "paragraph" ? `p:${paragraphIndex}` : `p:${paragraphIndex}:r:${runIndex}`;
  return asSourceNodeId(`text:${owner}:${suffix}`);
}
