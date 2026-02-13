import JSZip from "jszip";

export interface PptxArchive {
  files: Map<string, string>;
  media: Map<string, Uint8Array>;
}

export async function readPptx(input: Buffer | Uint8Array): Promise<PptxArchive> {
  const zip = await JSZip.loadAsync(input);

  const files = new Map<string, string>();
  const media = new Map<string, Uint8Array>();

  const promises: Promise<void>[] = [];

  zip.forEach((relativePath, zipEntry) => {
    if (zipEntry.dir) return;

    if (relativePath.startsWith("ppt/media/")) {
      promises.push(
        zipEntry.async("uint8array").then((buf) => {
          media.set(relativePath, buf);
        }),
      );
    } else if (
      relativePath.endsWith(".xml") ||
      relativePath.endsWith(".rels") ||
      relativePath === "[Content_Types].xml"
    ) {
      promises.push(
        zipEntry.async("string").then((str) => {
          files.set(relativePath, str);
        }),
      );
    }
  });

  await Promise.all(promises);
  return { files, media };
}
