import { strFromU8, unzipSync } from "fflate";

/**
 * メディアファイルの遅延読み込みマップ。
 * ZIP 内のメディアファイルを事前にすべて展開せず、
 * get() 呼び出し時に必要なファイルだけを解凍する。
 */
export class LazyMediaMap {
  private rawInput: Uint8Array;
  private cache = new Map<string, Uint8Array>();
  private entryIndex: Set<string>;

  constructor(rawInput: Uint8Array, mediaEntryNames: Set<string>) {
    this.rawInput = rawInput;
    this.entryIndex = mediaEntryNames;
  }

  get(path: string): Uint8Array | undefined {
    if (!this.entryIndex.has(path)) return undefined;

    const cached = this.cache.get(path);
    if (cached) return cached;

    const result = unzipSync(this.rawInput, {
      filter: (file) => file.name === path,
    });
    const data = result[path];
    if (data) {
      this.cache.set(path, data);
    }
    return data;
  }
}

export interface MediaAccessor {
  get(path: string): Uint8Array | undefined;
}

export interface PptxArchive {
  files: Map<string, string>;
  media: MediaAccessor;
}

export function readPptx(input: Buffer | Uint8Array): PptxArchive {
  const rawInput = new Uint8Array(input);
  const mediaEntryNames = new Set<string>();
  const unzipped = unzipSync(rawInput, {
    filter: (file) => {
      if (file.name.startsWith("ppt/media/")) {
        mediaEntryNames.add(file.name);
        return false;
      }
      return true;
    },
  });

  const files = new Map<string, string>();

  for (const [relativePath, data] of Object.entries(unzipped)) {
    if (relativePath.endsWith("/")) continue;

    if (
      relativePath.endsWith(".xml") ||
      relativePath.endsWith(".rels") ||
      relativePath === "[Content_Types].xml"
    ) {
      files.set(relativePath, strFromU8(data));
    }
  }

  const media = new LazyMediaMap(rawInput, mediaEntryNames);

  return { files, media };
}
