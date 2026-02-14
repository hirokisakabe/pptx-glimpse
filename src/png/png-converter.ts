import { Resvg } from "@resvg/resvg-wasm";
import { ensureWasmInitialized } from "./wasm-init.js";
import { tryLoadNativeResvg, type ResvgConstructor } from "./native-resvg.js";
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
  const NativeResvg = await tryLoadNativeResvg();

  if (NativeResvg) {
    return renderWithResvg(NativeResvg, svgString, options, true);
  }

  // WASM fallback
  await ensureWasmInitialized();
  return renderWithResvg(Resvg as unknown as ResvgConstructor, svgString, options, false);
}

function renderWithResvg(
  ResvgClass: ResvgConstructor,
  svgString: string,
  options: PngConvertOptions | undefined,
  defaultLoadSystemFonts: boolean,
): { png: Uint8Array; width: number; height: number } {
  const fitTo = options?.width
    ? { mode: "width" as const, value: options.width }
    : options?.height
      ? { mode: "height" as const, value: options.height }
      : { mode: "original" as const };

  const resvgOptions: Record<string, unknown> = { fitTo };

  const fontOption: Record<string, unknown> = {};
  if (options?.fonts) {
    const f = options.fonts;
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
  }

  // ネイティブ版: loadSystemFonts が未指定なら true をデフォルトに
  if (fontOption.loadSystemFonts === undefined && defaultLoadSystemFonts) {
    fontOption.loadSystemFonts = true;
  }

  if (Object.keys(fontOption).length > 0) {
    resvgOptions.font = fontOption;
  }

  const resvg = new ResvgClass(svgString, resvgOptions);

  const rendered = resvg.render();
  const png = rendered.asPng();
  const { width, height } = rendered;

  rendered.free?.();
  resvg.free?.();

  return { png, width, height };
}
