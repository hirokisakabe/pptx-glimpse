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

  if (options?.width || options?.height) {
    pipeline = pipeline.resize({
      width: options.width,
      height: options.height,
      fit: "contain",
      background: { r: 255, g: 255, b: 255, alpha: 1 },
    });
  }

  const result = await pipeline.png().toBuffer({ resolveWithObject: true });

  return {
    png: result.data,
    width: result.info.width,
    height: result.info.height,
  };
}
