import { Resvg, type ResvgRenderOptions } from "@resvg/resvg-wasm";

import type { PngConvertOptions, SvgToPngResult } from "./types.js";

export function renderSvgToPng(svgString: string, options?: PngConvertOptions): SvgToPngResult {
  const resvgOptions: ResvgRenderOptions = {};

  if (options?.width) {
    resvgOptions.fitTo = { mode: "width", value: options.width };
  } else if (options?.height) {
    resvgOptions.fitTo = { mode: "height", value: options.height };
  }

  const fontBuffers = options?.fontBuffers;
  if (fontBuffers && fontBuffers.length > 0) {
    resvgOptions.font = { fontBuffers };
  }

  const resvg = new Resvg(svgString, resvgOptions);
  const rendered = resvg.render();
  return {
    png: new Uint8Array(rendered.asPng()),
    width: rendered.width,
    height: rendered.height,
  };
}
