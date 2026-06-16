/**
 * テスト用: 複数の TTF ArrayBuffer から TTC バイナリを構築するヘルパー。
 *
 * 各 TTF 内のテーブルオフセットを TTC 全体の絶対アドレスに書き換える。
 */
export function buildTtcFromTtfs(ttfBuffers: ArrayBuffer[]): ArrayBuffer {
  const fontInfos: {
    sfVersion: number;
    tables: { tag: number; checkSum: number; offset: number; length: number }[];
  }[] = [];

  for (const buf of ttfBuffers) {
    const view = new DataView(buf);
    const sfVersion = view.getUint32(0);
    const numTables = view.getUint16(4);
    const tables: { tag: number; checkSum: number; offset: number; length: number }[] = [];
    for (let i = 0; i < numTables; i++) {
      const recOffset = 12 + i * 16;
      tables.push({
        tag: view.getUint32(recOffset),
        checkSum: view.getUint32(recOffset + 4),
        offset: view.getUint32(recOffset + 8),
        length: view.getUint32(recOffset + 12),
      });
    }
    fontInfos.push({ sfVersion, tables });
  }

  const ttcHeaderSize = 12 + ttfBuffers.length * 4;
  const fontHeaderSizes = fontInfos.map((info) => 12 + info.tables.length * 16);

  const fontOffsets: number[] = [];
  let pos = ttcHeaderSize;
  for (const size of fontHeaderSizes) {
    fontOffsets.push(pos);
    pos += size;
  }

  const dataStart = pos;
  const tableDataEntries: { srcBuf: ArrayBuffer; srcOffset: number; length: number }[] = [];
  const newTableOffsets: number[][] = [];
  let dataPos = dataStart;

  for (let fi = 0; fi < fontInfos.length; fi++) {
    const offsets: number[] = [];
    for (const table of fontInfos[fi].tables) {
      offsets.push(dataPos);
      tableDataEntries.push({
        srcBuf: ttfBuffers[fi],
        srcOffset: table.offset,
        length: table.length,
      });
      dataPos += (table.length + 3) & ~3;
    }
    newTableOffsets.push(offsets);
  }

  const totalSize = dataPos;
  const output = new ArrayBuffer(totalSize);
  const outView = new DataView(output);
  const outBytes = new Uint8Array(output);

  outView.setUint32(0, 0x74746366); // "ttcf"
  outView.setUint16(4, 1);
  outView.setUint16(6, 0);
  outView.setUint32(8, ttfBuffers.length);
  for (let i = 0; i < fontOffsets.length; i++) {
    outView.setUint32(12 + i * 4, fontOffsets[i]);
  }

  for (let fi = 0; fi < fontInfos.length; fi++) {
    const info = fontInfos[fi];
    const base = fontOffsets[fi];
    outView.setUint32(base, info.sfVersion);
    outView.setUint16(base + 4, info.tables.length);
    let searchRange = 16;
    let entrySelector = 0;
    while (searchRange * 2 <= info.tables.length * 16) {
      searchRange *= 2;
      entrySelector++;
    }
    outView.setUint16(base + 6, searchRange);
    outView.setUint16(base + 8, entrySelector);
    outView.setUint16(base + 10, info.tables.length * 16 - searchRange);

    for (let ti = 0; ti < info.tables.length; ti++) {
      const recOffset = base + 12 + ti * 16;
      outView.setUint32(recOffset, info.tables[ti].tag);
      outView.setUint32(recOffset + 4, info.tables[ti].checkSum);
      outView.setUint32(recOffset + 8, newTableOffsets[fi][ti]);
      outView.setUint32(recOffset + 12, info.tables[ti].length);
    }
  }

  let writePos = dataStart;
  for (const entry of tableDataEntries) {
    const src = new Uint8Array(entry.srcBuf, entry.srcOffset, entry.length);
    outBytes.set(src, writePos);
    writePos += (entry.length + 3) & ~3;
  }

  return output;
}
