import { readFileSync, existsSync, mkdirSync, writeFileSync } from "fs";
import { dirname } from "path";
import pixelmatch from "pixelmatch";
import sharp from "sharp";

export interface CompareResult {
  totalPixels: number;
  mismatchedPixels: number;
  mismatchPercentage: number;
  passed: boolean;
}

export interface CompareOptions {
  pixelThreshold: number;
  mismatchTolerance: number;
  /** 参照画像を actual のサイズにリサイズして比較する (LibreOffice VRT 用) */
  resizeRef?: boolean;
}

export async function compareImages(
  actualPng: Uint8Array | Buffer,
  referencePath: string,
  diffPath: string,
  options: CompareOptions,
): Promise<CompareResult> {
  if (!existsSync(referencePath)) {
    throw new Error(`Reference snapshot not found: ${referencePath}`);
  }

  const refPng = readFileSync(referencePath);

  const actualMeta = await sharp(actualPng).metadata();
  const width = actualMeta.width ?? 0;
  const height = actualMeta.height ?? 0;

  if (!options.resizeRef) {
    const refMeta = await sharp(refPng).metadata();
    if (width !== refMeta.width || height !== refMeta.height) {
      const total = width * height;
      return { totalPixels: total, mismatchedPixels: total, mismatchPercentage: 1, passed: false };
    }
  }

  const [actualRaw, refRaw] = await Promise.all([
    sharp(actualPng).ensureAlpha().raw().toBuffer(),
    options.resizeRef
      ? sharp(refPng).resize(width, height, { fit: "fill" }).ensureAlpha().raw().toBuffer()
      : sharp(refPng).ensureAlpha().raw().toBuffer(),
  ]);

  const totalPixels = width * height;
  const diffBuf = new Uint8Array(totalPixels * 4);

  const mismatched = pixelmatch(actualRaw, refRaw, diffBuf, width, height, {
    threshold: options.pixelThreshold,
    includeAA: false,
  });

  const mismatchPercentage = mismatched / totalPixels;
  const passed = mismatchPercentage <= options.mismatchTolerance;

  if (!passed) {
    mkdirSync(dirname(diffPath), { recursive: true });
    const diffPng = await sharp(Buffer.from(diffBuf), { raw: { width, height, channels: 4 } })
      .png()
      .toBuffer();
    writeFileSync(diffPath, diffPng);
  }

  return { totalPixels, mismatchedPixels: mismatched, mismatchPercentage, passed };
}
