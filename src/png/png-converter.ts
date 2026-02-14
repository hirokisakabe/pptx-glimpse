import sharp from "sharp";

export interface PngConvertOptions {
  width?: number;
  height?: number;
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<{ png: Buffer; width: number; height: number }> {
  const svgBuffer = Buffer.from(svgString);
  let pipeline = sharp(svgBuffer);

  if (options?.width) {
    pipeline = pipeline.resize(options.width);
  } else if (options?.height) {
    pipeline = pipeline.resize(null, options.height);
  }

  const result = await pipeline.png().toBuffer({ resolveWithObject: true });
  return { png: result.data, width: result.info.width, height: result.info.height };
}
