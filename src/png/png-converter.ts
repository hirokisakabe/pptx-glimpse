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
 * resvg-wasm の WASM モジュールを初期化する。
 * 明示的に呼び出さなくても、初回の PNG 変換時に自動的に初期化される。
 * アプリケーション起動時に初期化しておきたい場合に使用する。
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

  const resvg = new Resvg(svgString, resvgOptions);
  const rendered = resvg.render();
  return {
    png: Buffer.from(rendered.asPng()),
    width: rendered.width,
    height: rendered.height,
  };
}
