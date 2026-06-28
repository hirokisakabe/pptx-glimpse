import { initWasm, Resvg, type ResvgRenderOptions } from "@resvg/resvg-wasm";
import { readFile } from "fs/promises";
import { createRequire } from "module";

let wasmInitPromise: Promise<void> | null = null;

function resolveWasmPath(): string {
  // ESM: import.meta.url is available
  // CJS: tsup replaces import.meta with empty object, so fall back to __filename
  const baseUrl = import.meta.url || `file://${__filename}`;
  const require = createRequire(baseUrl);
  return require.resolve("@resvg/resvg-wasm/index_bg.wasm");
}

/**
 * Initializes the resvg-wasm WASM module.
 * Even when not called explicitly, it is initialized automatically on the first PNG conversion.
 * Use this when you want to initialize the application when it starts.
 */
export async function initResvgWasm(): Promise<void> {
  if (!wasmInitPromise) {
    wasmInitPromise = (async () => {
      const wasmPath = resolveWasmPath();
      const wasmBuffer = await readFile(wasmPath);
      await initWasm(wasmBuffer);
    })();
  }
  await wasmInitPromise;
}

interface PngConvertOptions {
  width?: number;
  height?: number;
  /** Font buffers used to render SVG <text> elements */
  fontBuffers?: Uint8Array[];
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<{ png: Buffer; width: number; height: number }> {
  await initResvgWasm();

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
    png: Buffer.from(rendered.asPng()),
    width: rendered.width,
    height: rendered.height,
  };
}
