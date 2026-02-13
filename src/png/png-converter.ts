import { Resvg } from "@resvg/resvg-wasm";
import { ensureWasmInitialized } from "./wasm-init.js";

export interface PngConvertOptions {
  width?: number;
  height?: number;
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<{ png: Uint8Array; width: number; height: number }> {
  await ensureWasmInitialized();

  const resvg = new Resvg(svgString, {
    fitTo: options?.width
      ? { mode: "width", value: options.width }
      : options?.height
        ? { mode: "height", value: options.height }
        : { mode: "original" },
  });

  const rendered = resvg.render();
  const png = rendered.asPng();
  const { width, height } = rendered;

  rendered.free();
  resvg.free();

  return { png, width, height };
}
