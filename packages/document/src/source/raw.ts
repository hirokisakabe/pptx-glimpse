/**
 * Raw OOXML escape hatch 関連の型 (`docs/raw-ooxml-round-trip.md`)。
 *
 * CleanDoc は supported semantics を typed field で表しつつ、未対応・部分対応の
 * XML を raw sidecar として、未編集の part を raw package part として保持し、
 * structural round-trip を成立させる。byte 一致は目標にしない。
 */

import type { PartPath, RawSidecarId } from "./handles.js";

/**
 * 部分的にパースされた raw OOXML ノード。namespace prefix 付きの qualified name、
 * 属性、子ノード、テキストを保持する。CleanDoc が typed に表現しない要素
 * (vendor extension / `mc:AlternateContent` / 未知の DrawingML 等) の保存に使う。
 */
export interface RawOoxmlNode {
  /** namespace prefix を含む要素名 (例: `a:extLst`)。 */
  readonly name: string;
  readonly attributes?: Readonly<Record<string, string>>;
  readonly children?: readonly RawOoxmlNode[];
  /** 要素の text content (持つ場合)。 */
  readonly text?: string;
}

/**
 * CleanDoc source node に付随する raw XML sidecar。最も近い source node に
 * 紐づけ、親要素内での順序メタデータを保持して書き戻し時の順序を復元する。
 */
export interface RawSidecar {
  readonly id: RawSidecarId;
  readonly node: RawOoxmlNode;
  /** 所有要素の子並びにおける順序スロット。 */
  readonly orderingSlot?: number;
}

/**
 * 未編集の package part をそのまま書き戻すための raw fallback。bytes (binary
 * asset) もしくは XML tree の **いずれか一方** で保持する判別共用体。両持ち /
 * 両欠の不正状態を型で排除する。
 */
export type RawPackagePart =
  | {
      readonly kind: "binary";
      readonly partPath: PartPath;
      readonly contentType: string;
      /** binary part (画像・埋め込みワークブック等) の元バイト列。 */
      readonly bytes: Uint8Array;
    }
  | {
      readonly kind: "xml";
      readonly partPath: PartPath;
      readonly contentType: string;
      /** XML part を tree として保持する場合の root ノード。 */
      readonly xml: RawOoxmlNode;
    };
