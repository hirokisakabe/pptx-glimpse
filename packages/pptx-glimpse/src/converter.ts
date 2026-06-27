import type { FontMapping } from "pptx-glimpse-renderer";
import type { LogLevel } from "pptx-glimpse-renderer";

import {
  convertPptxToPngViaDocumentPath,
  convertPptxToSvgViaDocumentPath,
} from "./experimental-document-renderer.js";

export interface ConvertOptions {
  /** 変換対象のスライド番号 (1始まり)。未指定で全スライド */
  slides?: number[];
  /** 出力画像の幅 (ピクセル)。デフォルト: 960 */
  width?: number;
  /** 出力画像の高さ (ピクセル)。widthと同時指定時はwidthが優先 */
  height?: number;
  /** 警告ログレベル。デフォルト: "off" */
  logLevel?: LogLevel;
  /** 追加のフォントディレクトリパス。システムフォントに加えて検索する */
  fontDirs?: string[];
  /** PPTX フォント名 → OSS 代替フォントのカスタムマッピング。デフォルトマッピングにマージされる */
  fontMapping?: FontMapping;
  /** true のとき OS のシステムフォントをスキャンせず fontDirs のみを使用する */
  skipSystemFonts?: boolean;
  /**
   * SVG でのテキスト出力方式。デフォルト: "path"
   * - "path": グリフをアウトライン化した <path> として出力する。フォント環境に依存しない
   * - "text": ネイティブ <text> 要素 + サブセット化フォントの @font-face (data URI) 埋め込みで出力する。
   *   ブラウザでのインライン表示時にネイティブテキスト描画 (ヒンティング等) が効き、テキスト選択も可能になる。
   *   <img src="...svg"> 参照やサニタイズ環境では期待どおり描画されないことがある。
   *   convertPptxToPng では無視され、常に "path" で変換される (resvg は @font-face を解釈しないため)
   */
  textOutput?: "path" | "text";
}

export interface SlideSvg {
  slideNumber: number;
  svg: string;
}

export interface SlideImage {
  slideNumber: number;
  png: Buffer;
  width: number;
  height: number;
}

export async function convertPptxToSvg(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideSvg[]> {
  const result = await convertPptxToSvgViaDocumentPath(input, options);
  return [...result.slides];
}

export async function convertPptxToPng(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  const result = await convertPptxToPngViaDocumentPath(input, options);
  return [...result.slides];
}
