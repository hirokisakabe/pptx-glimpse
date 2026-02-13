import { Resvg } from "@resvg/resvg-wasm";
import { ensureWasmInitialized } from "./wasm-init.js";
import type { FontOptions } from "../converter.js";

export interface PngConvertOptions {
  width?: number;
  height?: number;
  fonts?: FontOptions;
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<{ png: Uint8Array; width: number; height: number }> {
  await ensureWasmInitialized();

  const fitTo = options?.width
    ? { mode: "width" as const, value: options.width }
    : options?.height
      ? { mode: "height" as const, value: options.height }
      : { mode: "original" as const };

  const resvgOptions: Record<string, unknown> = { fitTo };

  if (options?.fonts) {
    const f = options.fonts;
    const fontOption: Record<string, unknown> = {};
    if (f.loadSystemFonts !== undefined) fontOption.loadSystemFonts = f.loadSystemFonts;
    if (f.fontFiles) fontOption.fontFiles = f.fontFiles;
    if (f.fontDirs) fontOption.fontDirs = f.fontDirs;
    if (f.fontBuffers) {
      fontOption.fontBuffers = f.fontBuffers.map((b) =>
        b.data instanceof Uint8Array ? b.data : new Uint8Array(b.data),
      );
    }
    if (f.defaultFontFamily) fontOption.defaultFontFamily = f.defaultFontFamily;
    if (f.sansSerifFamily) fontOption.sansSerifFamily = f.sansSerifFamily;
    if (f.serifFamily) fontOption.serifFamily = f.serifFamily;
    resvgOptions.font = fontOption;
  }

  const resvg = new Resvg(svgString, resvgOptions);

  const rendered = resvg.render();
  const png = rendered.asPng();
  const { width, height } = rendered;

  rendered.free();
  resvg.free();

  return { png, width, height };
}
