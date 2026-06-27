/**
 * Raw OOXML escape hatch のための変換ヘルパー。
 *
 * CleanDoc が typed に表現しない要素 (未対応の fill / effect / extension 等) を
 * `RawOoxmlNode` / `RawSidecar` として保存し、structural round-trip を成立させる
 * (`docs/raw-ooxml-round-trip.md`)。fast-xml-parser が生成した plain object を
 * qualified name を保ったまま raw tree へ写し取る。
 */

import type { RawOoxmlNode, RawSidecar, RawSidecarId } from "../source/index.js";
import { asRawSidecarId } from "../source/index.js";
import { localName, type XmlNode } from "./xml.js";

const ATTR_PREFIX = "@_";
const TEXT_KEY = "#text";

/** part ごとに安定した raw sidecar id を発番する factory を作る。 */
export function createSidecarIdFactory(partPath: string): () => RawSidecarId {
  let counter = 0;
  return () => asRawSidecarId(`${partPath}#raw-${counter++}`);
}

/**
 * fast-xml-parser が生成した要素値を `RawOoxmlNode` に変換する。`name` は
 * qualified name (例: `a:effectLst`)。属性は `@_` prefix を除去して保持し、
 * `#text` は text content として保持する。
 */
function xmlValueToRawNode(name: string, value: unknown): RawOoxmlNode {
  if (typeof value !== "object" || value === null) {
    const text = scalarText(value);
    return text !== undefined && text !== "" ? { name, text } : { name };
  }

  const obj = value as Record<string, unknown>;
  const attributes: Record<string, string> = {};
  const children: RawOoxmlNode[] = [];
  let text: string | undefined;

  for (const key of Object.keys(obj)) {
    if (key === TEXT_KEY) {
      text = scalarText(obj[key]);
      continue;
    }
    if (key.startsWith(ATTR_PREFIX)) {
      const attrValue = scalarText(obj[key]);
      if (attrValue !== undefined) attributes[key.slice(ATTR_PREFIX.length)] = attrValue;
      continue;
    }
    const childValue = obj[key];
    const items = Array.isArray(childValue) ? childValue : [childValue];
    for (const item of items) {
      children.push(xmlValueToRawNode(key, item));
    }
  }

  return {
    name,
    ...(Object.keys(attributes).length > 0 ? { attributes } : {}),
    ...(children.length > 0 ? { children } : {}),
    ...(text !== undefined && text !== "" ? { text } : {}),
  };
}

/** 単一の子要素 (qualified `name` / value) から raw sidecar を 1 つ作る。 */
export function makeSidecar(
  name: string,
  value: unknown,
  nextId: () => RawSidecarId,
  orderingSlot?: number,
): RawSidecar {
  return {
    id: nextId(),
    node: xmlValueToRawNode(name, value),
    ...(orderingSlot !== undefined ? { orderingSlot } : {}),
  };
}

/**
 * 親要素の子のうち、`knownLocalNames` に含まれない (= typed に解釈しない) もの
 * すべてを raw sidecar として収集する。属性キー・テキストキーは対象外。
 * 子の出現順を `orderingSlot` として記録し、書き戻し時の順序復元に備える。
 */
export function collectUnknownSidecars(
  parent: XmlNode | undefined,
  knownLocalNames: ReadonlySet<string>,
  nextId: () => RawSidecarId,
): RawSidecar[] {
  if (!parent) return [];
  const sidecars: RawSidecar[] = [];
  let slot = 0;
  for (const key of Object.keys(parent)) {
    if (key.startsWith(ATTR_PREFIX) || key === TEXT_KEY) continue;
    const value = parent[key];
    const items = Array.isArray(value) ? value : [value];
    if (knownLocalNames.has(localName(key))) {
      slot += items.length;
      continue;
    }
    for (const item of items) {
      sidecars.push(makeSidecar(key, item, nextId, slot));
      slot++;
    }
  }
  return sidecars;
}

function scalarText(value: unknown): string | undefined {
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return undefined;
}
