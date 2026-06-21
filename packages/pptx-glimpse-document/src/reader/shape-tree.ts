/**
 * `p:spTree` を CleanDoc source の shape node 列へ読み取る。
 *
 * simple autoshape (`p:sp`) と embedded raster image (`p:pic`) を typed に表し、
 * group / connector / graphicFrame 等の未対応ノードは raw shape node として
 * 保存する (`docs/cleandoc-minimal-poc-scope.md`)。typed node 内でも、未対応の
 * 子要素・属性は raw sidecar として保持する。
 *
 * 注意: fast-xml-parser は同名要素のみを配列化するため、子の反復順は
 * 「タグ種別の初出順 × 同名内の文書順」となり、異種タグ間の完全な文書順
 * (z-order) は保証されない。effective element ordering の解決は computed view
 * (#465) の責務であり、PoC 対象 fixture は sp/pic が交互配置されないため
 * 実害はない。
 */

import type {
  PartPath,
  RawSidecarId,
  SourceImage,
  SourceImageCrop,
  SourceNodeId,
  SourcePlaceholder,
  SourceRawShapeNode,
  SourceShape,
  SourceShapeNode,
} from "../source/index.js";
import { asOoxmlPercent, asRelationshipId, asSourceNodeId } from "../source/index.js";
import { numericAttr, parseFill, parseOutline, parseTransform } from "./drawing.js";
import { collectUnknownSidecars, makeSidecar } from "./raw-node.js";
import { parseTextBody } from "./text.js";
import { getAttr, getChild, getNamespacedAttr, localName, type XmlNode } from "./xml.js";

const KNOWN_SHAPE_CHILDREN: ReadonlySet<string> = new Set(["nvSpPr", "spPr", "txBody"]);
const KNOWN_PICTURE_CHILDREN: ReadonlySet<string> = new Set(["nvPicPr", "blipFill", "spPr"]);
const KNOWN_BLIP_FILL_CHILDREN: ReadonlySet<string> = new Set(["blip", "srcRect"]);
// `a:spPr` のうち typed に解釈する子。これ以外 (custGeom / effectLst / scene3d /
// extLst 等) は raw sidecar として保持する。fill 系は parseFill が typed/raw を
// 判別するため known 扱いにして二重計上を防ぐ。
const KNOWN_SP_PR_CHILDREN: ReadonlySet<string> = new Set([
  "xfrm",
  "prstGeom",
  "solidFill",
  "noFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
  "ln",
]);

/** `p:spTree` を読み、shape node 列を返す。 */
export function parseShapeTree(
  spTree: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
): SourceShapeNode[] {
  if (!spTree) return [];

  const nodes: SourceShapeNode[] = [];
  let orderingSlot = 0;
  for (const key of Object.keys(spTree)) {
    if (key.startsWith("@_")) continue;
    const local = localName(key);
    // group 自身の非可視プロパティはノードではないため除外する。
    if (local === "nvGrpSpPr" || local === "grpSpPr") continue;

    const value = spTree[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      const node = item as XmlNode;
      if (local === "sp") {
        nodes.push(parseShape(node, partPath, nextId, orderingSlot));
      } else if (local === "pic") {
        nodes.push(parseImage(node, partPath, nextId, orderingSlot));
      } else {
        nodes.push(parseRawShapeNode(key, node, partPath, nextId, orderingSlot));
      }
      orderingSlot++;
    }
  }
  return nodes;
}

function parseShape(
  sp: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceShape {
  const nvSpPr = getChild(sp, "nvSpPr");
  const cNvPr = getChild(nvSpPr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const placeholder = parsePlaceholder(getChild(getChild(nvSpPr, "nvPr"), "ph"));

  const spPr = getChild(sp, "spPr");
  const transform = parseTransform(spPr);
  const geometry = parsePresetGeometry(spPr);
  const fill = parseFill(spPr, nextId);
  const outline = parseOutline(spPr, nextId);
  const textBody = parseTextBody(getChild(sp, "txBody"), partPath, nextId, nodeId, orderingSlot);

  const rawSidecars = [
    ...collectUnknownSidecars(sp, KNOWN_SHAPE_CHILDREN, nextId),
    ...collectUnknownSidecars(spPr, KNOWN_SP_PR_CHILDREN, nextId),
  ];

  return {
    kind: "shape",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(geometry !== undefined ? { geometry } : {}),
    ...(fill !== undefined ? { fill } : {}),
    ...(outline !== undefined ? { outline } : {}),
    ...(textBody !== undefined ? { textBody } : {}),
    ...(placeholder !== undefined ? { placeholder } : {}),
    handle: { partPath, ...(nodeId !== undefined ? { nodeId } : {}), orderingSlot },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseImage(
  pic: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceImage {
  const cNvPr = getChild(getChild(pic, "nvPicPr"), "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");

  const blipFill = getChild(pic, "blipFill");
  const blip = getChild(blipFill, "blip");
  const embed = getNamespacedAttr(blip, "embed");
  const crop = parseCrop(getChild(blipFill, "srcRect"));

  const spPr = getChild(pic, "spPr");
  const transform = parseTransform(spPr);

  const rawSidecars = [
    ...collectUnknownSidecars(pic, KNOWN_PICTURE_CHILDREN, nextId),
    ...collectUnknownSidecars(spPr, KNOWN_SP_PR_CHILDREN, nextId),
    // `a:stretch` / `a:tile` 等の fill mode と blip 配下の recolor 操作を保持する。
    ...collectUnknownSidecars(blipFill, KNOWN_BLIP_FILL_CHILDREN, nextId),
    ...collectUnknownSidecars(blip, EMPTY_KNOWN, nextId),
  ];

  return {
    kind: "image",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(embed !== undefined ? { blipRelationshipId: asRelationshipId(embed) } : {}),
    ...(crop !== undefined ? { crop } : {}),
    handle: {
      partPath,
      ...(nodeId !== undefined ? { nodeId } : {}),
      ...(embed !== undefined ? { relationshipId: asRelationshipId(embed) } : {}),
      orderingSlot,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseRawShapeNode(
  key: string,
  node: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceRawShapeNode {
  const nodeId = sourceNodeId(getChild(getChild(node, "nvSpPr"), "cNvPr"));
  return {
    kind: "raw",
    ...(nodeId !== undefined ? { nodeId } : {}),
    raw: makeSidecar(key, node, nextId),
    handle: { partPath, ...(nodeId !== undefined ? { nodeId } : {}), orderingSlot },
  };
}

function parsePlaceholder(ph: XmlNode | undefined): SourcePlaceholder | undefined {
  if (!ph) return undefined;
  const type = getAttr(ph, "type");
  const index = numericAttr(ph, "idx");
  if (type === undefined && index === undefined) return undefined;
  return {
    ...(type !== undefined ? { type } : {}),
    ...(index !== undefined ? { index } : {}),
  };
}

function parsePresetGeometry(spPr: XmlNode | undefined) {
  const prstGeom = getChild(spPr, "prstGeom");
  if (!prstGeom) return undefined;
  const preset = getAttr(prstGeom, "prst");
  return preset !== undefined ? { preset } : undefined;
}

function parseCrop(srcRect: XmlNode | undefined): SourceImageCrop | undefined {
  if (!srcRect) return undefined;
  const left = numericAttr(srcRect, "l");
  const top = numericAttr(srcRect, "t");
  const right = numericAttr(srcRect, "r");
  const bottom = numericAttr(srcRect, "b");
  const crop: SourceImageCrop = {
    ...(left !== undefined ? { left: asOoxmlPercent(left) } : {}),
    ...(top !== undefined ? { top: asOoxmlPercent(top) } : {}),
    ...(right !== undefined ? { right: asOoxmlPercent(right) } : {}),
    ...(bottom !== undefined ? { bottom: asOoxmlPercent(bottom) } : {}),
  };
  return Object.keys(crop).length > 0 ? crop : undefined;
}

function sourceNodeId(cNvPr: XmlNode | undefined): SourceNodeId | undefined {
  const id = getAttr(cNvPr, "id");
  return id !== undefined ? asSourceNodeId(id) : undefined;
}

const EMPTY_KNOWN: ReadonlySet<string> = new Set();
