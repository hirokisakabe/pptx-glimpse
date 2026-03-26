import { renderAsync, type ResvgRenderOptions } from "@resvg/resvg-js";

interface PngConvertOptions {
  width?: number;
  height?: number;
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<{ png: Buffer; width: number; height: number }> {
  const resvgOptions: ResvgRenderOptions = {};

  if (options?.width) {
    resvgOptions.fitTo = { mode: "width", value: options.width };
  } else if (options?.height) {
    resvgOptions.fitTo = { mode: "height", value: options.height };
  }

  const rendered = await renderAsync(svgString, resvgOptions);
  return {
    png: Buffer.from(rendered.asPng()),
    width: rendered.width,
    height: rendered.height,
  };
}
