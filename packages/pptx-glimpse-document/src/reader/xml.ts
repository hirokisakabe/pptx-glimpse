/**
 * reader 内部で使う最小限の OOXML XML ヘルパー。
 *
 * `@pptx-glimpse/document` は下位基盤であり renderer 側の XML パーサーを参照
 * できないため、fast-xml-parser を直接利用する。namespace prefix は **保持**
 * する (`removeNSPrefix: false`)。これは `p:sldId` が namespace 無しの `id`
 * (スライド ID) と relationships namespace の `r:id` (relationship 参照) を
 * 同時に持ち、prefix を落とすと両者が衝突して relationship 参照を復元できなく
 * なるため。要素アクセスは local name (prefix を無視) で行い、属性は plain /
 * namespaced を区別して取得する。
 */

import { XMLParser } from "fast-xml-parser";

export type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseAttributeValue: false,
  // prefix を保持する。理由はファイル冒頭コメント参照。
  removeNSPrefix: false,
  trimValues: true,
});

/** XML 文字列をパースして root オブジェクトを返す。 */
export function parseXml(xml: string): XmlNode {
  return parser.parse(xml) as XmlNode;
}

/** `a:foo` のような qualified name から local part (`foo`) を取り出す。 */
function localName(key: string): string {
  const colon = key.indexOf(":");
  return colon === -1 ? key : key.slice(colon + 1);
}

/**
 * 子要素を local name で取得する (prefix は無視)。属性キー (`@_`) は対象外。
 * 同名要素が複数ある場合は最初の一致を返す。
 */
export function getChild(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === name) {
      const value = node[key];
      return Array.isArray(value)
        ? (value[0] as XmlNode | undefined)
        : (value as XmlNode | undefined);
    }
  }
  return undefined;
}

/** 子要素を local name で取得し、常に配列として返す。 */
export function getChildArray(node: XmlNode | undefined, name: string): XmlNode[] {
  if (!node) return [];
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === name) {
      const value = node[key];
      if (value === undefined || value === null) return [];
      return (Array.isArray(value) ? value : [value]) as XmlNode[];
    }
  }
  return [];
}

/** namespace 無しの属性 (`@_<name>`) を取得する。 */
export function getAttr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  return scalarToString(node[`@_${name}`]);
}

/**
 * namespace 付き属性 (`@_<prefix>:<localName>`) を取得する。`p:sldId` の
 * `r:id` のように、plain な `id` と区別して relationship 参照を取り出すために
 * 使う。prefix は問わず、local part が一致する最初の属性を返す。
 */
export function getNamespacedAttr(
  node: XmlNode | undefined,
  localAttr: string,
): string | undefined {
  if (!node) return undefined;
  for (const key of Object.keys(node)) {
    if (!key.startsWith("@_")) continue;
    const attr = key.slice(2);
    const colon = attr.indexOf(":");
    if (colon !== -1 && attr.slice(colon + 1) === localAttr) {
      const value = scalarToString(node[key]);
      if (value !== undefined) return value;
    }
  }
  return undefined;
}

/** 属性値 (string/number/boolean) を文字列化する。object 等は undefined。 */
function scalarToString(value: unknown): string | undefined {
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return undefined;
}
