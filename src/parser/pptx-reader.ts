import { unzipSync, strFromU8 } from "fflate";

export interface PptxArchive {
  files: Map<string, string>;
  media: Map<string, Uint8Array>;
}

export function readPptx(input: Buffer | Uint8Array): PptxArchive {
  const unzipped = unzipSync(new Uint8Array(input));

  const files = new Map<string, string>();
  const media = new Map<string, Uint8Array>();

  for (const [relativePath, data] of Object.entries(unzipped)) {
    if (relativePath.endsWith("/")) continue;

    if (relativePath.startsWith("ppt/media/")) {
      media.set(relativePath, data);
    } else if (
      relativePath.endsWith(".xml") ||
      relativePath.endsWith(".rels") ||
      relativePath === "[Content_Types].xml"
    ) {
      files.set(relativePath, strFromU8(data));
    }
  }

  return { files, media };
}
