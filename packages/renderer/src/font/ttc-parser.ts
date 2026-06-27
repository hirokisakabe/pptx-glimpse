/**
 * TTC (TrueType Collection) バイナリの解析・分割モジュール。
 *
 * TTC ファイルから個別の TTF/OTF バッファを抽出し、
 * opentype.js でパースできる形式にする。
 */

const TTC_TAG = 0x74746366; // "ttcf"

/**
 * バッファが TTC (TrueType Collection) 形式かどうかを判定する。
 * 先頭 4 バイトが "ttcf" であれば TTC と判定する。
 */
export function isTtcBuffer(data: ArrayBuffer | Uint8Array): boolean {
  const view = toDataView(data);
  if (view.byteLength < 4) return false;
  return view.getUint32(0) === TTC_TAG;
}

/**
 * TTC バッファから個別の TTF/OTF バッファを抽出する。
 *
 * 各フォントの OffsetTable + テーブルデータを独立した TTF/OTF バッファに再パックする。
 * TTC 内でテーブルが共有されている場合、各フォントがそれぞれコピーを持つことになる。
 *
 * @returns 抽出されたフォントバッファの配列。TTC でないか解析に失敗した場合は空配列。
 */
export function extractTtcFonts(data: ArrayBuffer | Uint8Array): ArrayBuffer[] {
  const view = toDataView(data);
  const bytes = toUint8Array(data);

  if (view.byteLength < 12) return [];
  if (view.getUint32(0) !== TTC_TAG) return [];

  const numFonts = view.getUint32(8);
  if (numFonts === 0) return [];

  const headerEnd = 12 + numFonts * 4;
  if (view.byteLength < headerEnd) return [];

  const results: ArrayBuffer[] = [];

  for (let i = 0; i < numFonts; i++) {
    try {
      const fontOffset = view.getUint32(12 + i * 4);
      const extracted = extractSingleFont(view, bytes, fontOffset);
      if (extracted) results.push(extracted);
    } catch {
      // 個別フォントの抽出失敗はスキップ
    }
  }

  return results;
}

/**
 * TTC 内の単一フォントを独立した TTF/OTF バッファとして再パックする。
 */
function extractSingleFont(
  view: DataView,
  bytes: Uint8Array,
  fontOffset: number,
): ArrayBuffer | null {
  if (fontOffset + 12 > view.byteLength) return null;

  const sfVersion = view.getUint32(fontOffset);
  const numTables = view.getUint16(fontOffset + 4);

  if (numTables === 0) return null;

  const tableRecordsStart = fontOffset + 12;
  const tableRecordsEnd = tableRecordsStart + numTables * 16;
  if (tableRecordsEnd > view.byteLength) return null;

  // テーブル情報を読み取り、境界チェック
  const tables: { tag: number; checkSum: number; offset: number; length: number }[] = [];
  for (let i = 0; i < numTables; i++) {
    const recOffset = tableRecordsStart + i * 16;
    const tableOffset = view.getUint32(recOffset + 8);
    const tableLength = view.getUint32(recOffset + 12);

    // テーブルデータがバッファ範囲外の場合はフォント全体を無効とする
    if (tableOffset > view.byteLength || tableLength > view.byteLength - tableOffset) {
      return null;
    }

    tables.push({
      tag: view.getUint32(recOffset),
      checkSum: view.getUint32(recOffset + 4),
      offset: tableOffset,
      length: tableLength,
    });
  }

  // 出力サイズを計算
  const headerSize = 12 + numTables * 16;
  let dataSize = 0;
  for (const table of tables) {
    dataSize += alignTo4(table.length);
  }
  const totalSize = headerSize + dataSize;

  // 出力バッファを構築
  const output = new ArrayBuffer(totalSize);
  const outView = new DataView(output);
  const outBytes = new Uint8Array(output);

  // OffsetTable ヘッダーを書き込み
  outView.setUint32(0, sfVersion);
  outView.setUint16(4, numTables);

  // searchRange, entrySelector, rangeShift を計算
  const { searchRange, entrySelector, rangeShift } = calcOffsetTableFields(numTables);
  outView.setUint16(6, searchRange);
  outView.setUint16(8, entrySelector);
  outView.setUint16(10, rangeShift);

  // テーブルレコードとデータを書き込み
  let currentDataOffset = headerSize;
  for (let i = 0; i < numTables; i++) {
    const table = tables[i];
    const recOffset = 12 + i * 16;

    // テーブルレコード
    outView.setUint32(recOffset, table.tag);
    outView.setUint32(recOffset + 4, table.checkSum);
    outView.setUint32(recOffset + 8, currentDataOffset);
    outView.setUint32(recOffset + 12, table.length);

    // テーブルデータをコピー
    outBytes.set(bytes.subarray(table.offset, table.offset + table.length), currentDataOffset);

    currentDataOffset += alignTo4(table.length);
  }

  return output;
}

/** 4 バイト境界にアラインする */
function alignTo4(n: number): number {
  return (n + 3) & ~3;
}

/** OffsetTable の searchRange, entrySelector, rangeShift を計算する */
function calcOffsetTableFields(numTables: number): {
  searchRange: number;
  entrySelector: number;
  rangeShift: number;
} {
  let searchRange = 16;
  let entrySelector = 0;
  while (searchRange * 2 <= numTables * 16) {
    searchRange *= 2;
    entrySelector++;
  }
  const rangeShift = numTables * 16 - searchRange;
  return { searchRange, entrySelector, rangeShift };
}

function toDataView(data: ArrayBuffer | Uint8Array): DataView {
  if (data instanceof Uint8Array) {
    return new DataView(data.buffer, data.byteOffset, data.byteLength);
  }
  return new DataView(data);
}

function toUint8Array(data: ArrayBuffer | Uint8Array): Uint8Array {
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data);
}
