/**
 * Module for parsing and splitting TTC (TrueType Collection) binaries.
 *
 * Extracts individual TTF/OTF buffers from a TTC file
 * opentype.js and makes them parseable by opentype.js.
 */

const TTC_TAG = 0x74746366; // "ttcf"

/**
 * Determines whether the buffer is in TTC (TrueType Collection) format.
 * The first 4 bytes "ttcf" indicate TTC.
 */
export function isTtcBuffer(data: ArrayBuffer | Uint8Array): boolean {
  const view = toDataView(data);
  if (view.byteLength < 4) return false;
  return view.getUint32(0) === TTC_TAG;
}

/**
 * Extracts individual TTF/OTF buffers from a TTC buffer.
 *
 * Repacks each font's OffsetTable + table data into an independent TTF/OTF buffer.
 * If the table is shared within the TTC, each font will have its own copy.
 *
 * @returns Array of extracted font buffers. Empty array if not TTC or parsing failed.
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
      // Skip failures when extracting individual fonts
    }
  }

  return results;
}

/**
 * Repack a single font in TTC as a separate TTF/OTF buffer.
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

  // Read table information and check bounds
  const tables: { tag: number; checkSum: number; offset: number; length: number }[] = [];
  for (let i = 0; i < numTables; i++) {
    const recOffset = tableRecordsStart + i * 16;
    const tableOffset = view.getUint32(recOffset + 8);
    const tableLength = view.getUint32(recOffset + 12);

    // If table data is outside the buffer range, invalidate the entire font
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

  // Calculate output size
  const headerSize = 12 + numTables * 16;
  let dataSize = 0;
  for (const table of tables) {
    dataSize += alignTo4(table.length);
  }
  const totalSize = headerSize + dataSize;

  // Build output buffer
  const output = new ArrayBuffer(totalSize);
  const outView = new DataView(output);
  const outBytes = new Uint8Array(output);

  // OffsetTable Write header
  outView.setUint32(0, sfVersion);
  outView.setUint16(4, numTables);

  // searchRange, entrySelector, rangeShift calculate
  const { searchRange, entrySelector, rangeShift } = calcOffsetTableFields(numTables);
  outView.setUint16(6, searchRange);
  outView.setUint16(8, entrySelector);
  outView.setUint16(10, rangeShift);

  // Write table records and data
  let currentDataOffset = headerSize;
  for (let i = 0; i < numTables; i++) {
    const table = tables[i];
    const recOffset = 12 + i * 16;

    // Table record
    outView.setUint32(recOffset, table.tag);
    outView.setUint32(recOffset + 4, table.checkSum);
    outView.setUint32(recOffset + 8, currentDataOffset);
    outView.setUint32(recOffset + 12, table.length);

    // Copy table data
    outBytes.set(bytes.subarray(table.offset, table.offset + table.length), currentDataOffset);

    currentDataOffset += alignTo4(table.length);
  }

  return output;
}

/** Aligns to a 4-byte boundary */
function alignTo4(n: number): number {
  return (n + 3) & ~3;
}

/** Calculate searchRange, entrySelector, rangeShift of OffsetTable */
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
