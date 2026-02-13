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
]);

export function createXmlParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    removeNSPrefix: true,
    htmlEntities: true,
    isArray: (_name: string, jpath: string) => {
      const tag = jpath.split(".").pop() ?? "";
      return ARRAY_TAGS.has(tag);
    },
  });
}

export function parseXml(xml: string): Record<string, unknown> {
  const parser = createXmlParser();
  return parser.parse(xml) as Record<string, unknown>;
}

// preserveOrder: true で子要素の出現順序を保持するパーサー。
// spTree 内の異なる要素タイプ（sp, pic, cxnSp 等）の Z-order を正しく復元するために使用。
// データ取得には既存の parseXml を使い、本関数は順序情報のみに使用する。
export function parseXmlOrdered(xml: string): XmlOrderedNode[] {
  const parser = new XMLParser({
    preserveOrder: true,
    removeNSPrefix: true,
    ignoreAttributes: true,
  });
  return parser.parse(xml) as XmlOrderedNode[];
}
