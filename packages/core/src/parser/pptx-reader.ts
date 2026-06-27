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

/**
 * XML/RELS ファイルの遅延読み込みマップ。
 * slides オプションで一部スライドのみ要求された場合に、
 * 不要なスライド XML の解凍を省く。
 */
export class LazyXmlMap {
  private rawInput: Uint8Array;
  private cache = new Map<string, string>();
  private entryIndex: Set<string>;

  constructor(rawInput: Uint8Array, xmlEntryNames: Set<string>) {
    this.rawInput = rawInput;
    this.entryIndex = xmlEntryNames;
  }

  has(path: string): boolean {
    return this.entryIndex.has(path);
  }

  get(path: string): string | undefined {
    if (!this.entryIndex.has(path)) return undefined;

    const cached = this.cache.get(path);
    if (cached !== undefined) return cached;

    const result = unzipSync(this.rawInput, {
      filter: (file) => file.name === path,
    });
    const data = result[path];
    if (data) {
      const str = strFromU8(data);
      this.cache.set(path, str);
      return str;
    }
    return undefined;
  }
}

export interface XmlAccessor {
  get(path: string): string | undefined;
  has(path: string): boolean;
}

export interface PptxArchive {
  files: XmlAccessor;
  media: MediaAccessor;
}

export function readPptx(input: Buffer | Uint8Array): PptxArchive {
  const rawInput = new Uint8Array(input);
  const xmlEntryNames = new Set<string>();
  const mediaEntryNames = new Set<string>();

  unzipSync(rawInput, {
    filter: (file) => {
      if (file.name.startsWith("ppt/media/")) {
        mediaEntryNames.add(file.name);
      } else if (
        file.name.endsWith(".xml") ||
        file.name.endsWith(".rels") ||
        file.name === "[Content_Types].xml"
      ) {
        xmlEntryNames.add(file.name);
      }
      return false;
    },
  });

  const files = new LazyXmlMap(rawInput, xmlEntryNames);
  const media = new LazyMediaMap(rawInput, mediaEntryNames);

  return { files, media };
}
