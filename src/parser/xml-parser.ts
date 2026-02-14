import { XMLParser } from "fast-xml-parser";

/** fast-xml-parser が返す XML ノードの型エイリアス */
export type XmlNode = Record<string, unknown>;

/** preserveOrder: true で返される順序付き XML ノード */
export type XmlOrderedNode = Record<string, unknown>;

// OOXML XML で単一要素でも配列として扱う必要があるタグ。
// fast-xml-parser は子要素が 1 つだとオブジェクト、複数だと配列を返すため、
// スライド上に図形が 1 つだけの場合などにパース結果が不安定になる。
// isArray で常に配列化することで、下流コードを統一的に記述できる。
const ARRAY_TAGS = new Set([
  "sp", // 図形 (Shape)
  "pic", // 画像 (Picture)
  "cxnSp", // コネクタ (Connector)
  "grpSp", // グループ (Group Shape)
  "graphicFrame", // テーブル・チャート等のフレーム
  "p", // テキスト段落 (Paragraph)
  "r", // テキストラン (Run)
  "br", // 改行 (Break)
  "fld", // フィールドコード (Field)
  "Relationship", // リレーションシップ
  "sldId", // スライド ID
  "gs", // グラデーションストップ (Gradient Stop)
  "gridCol", // テーブル列定義
  "tr", // テーブル行 (Table Row)
  "tc", // テーブルセル (Table Cell)
  "ser", // チャートデータ系列 (Series)
  "pt", // チャートデータポイント (Point)
  "gd", // ガイド定義 (Guide Definition)
  "ds", // カスタムダッシュセグメント (Custom Dash Segment)
  "AlternateContent", // mc:AlternateContent (SmartArt 等)
  "embeddedFont", // 埋め込みフォント (Embedded Font)
  "effectStyle", // エフェクトスタイル (Effect Style)
]);

// シングルトンパーサーインスタンス。
// XMLParser.parse() はステートレスなため、安全に再利用できる。
const standardParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  htmlEntities: true,
  isArray: (_name: string, jpath: string) => {
    const tag = jpath.split(".").pop() ?? "";
    return ARRAY_TAGS.has(tag);
  },
});

const orderedParser = new XMLParser({
  preserveOrder: true,
  removeNSPrefix: true,
  ignoreAttributes: true,
});

// 変換単位のパース結果キャッシュ。
// 同一 XML 文字列（マスター・レイアウト等）の重複パースを回避する。
let xmlCache: Map<string, Record<string, unknown>> | null = null;
let xmlOrderedCache: Map<string, XmlOrderedNode[]> | null = null;

/** 変換開始時にキャッシュを有効化する */
export function enableXmlCache(): void {
  xmlCache = new Map();
  xmlOrderedCache = new Map();
}

/** 変換完了時にキャッシュをクリアする（メモリリーク防止） */
export function clearXmlCache(): void {
  xmlCache = null;
  xmlOrderedCache = null;
}

export function parseXml(xml: string): Record<string, unknown> {
  if (xmlCache) {
    const cached = xmlCache.get(xml);
    if (cached) return cached;
  }
  const result = standardParser.parse(xml) as Record<string, unknown>;
  xmlCache?.set(xml, result);
  return result;
}

// preserveOrder: true で子要素の出現順序を保持するパーサー。
// spTree 内の異なる要素タイプ（sp, pic, cxnSp 等）の Z-order を正しく復元するために使用。
// データ取得には既存の parseXml を使い、本関数は順序情報のみに使用する。
export function parseXmlOrdered(xml: string): XmlOrderedNode[] {
  if (xmlOrderedCache) {
    const cached = xmlOrderedCache.get(xml);
    if (cached) return cached;
  }
  const result = orderedParser.parse(xml) as XmlOrderedNode[];
  xmlOrderedCache?.set(xml, result);
  return result;
}
