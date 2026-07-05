import { initWasm } from "@resvg/resvg-wasm";

import { renderSvgToPng } from "./render-svg.js";
import type { PngConvertOptions, ResvgWasmInput, SvgToPngResult } from "./types.js";

let wasmInitPromise: Promise<void> | null = null;

const nodeImportPrefix = "node:";
const fsModuleName = "fs";
const fsPromisesSegment = "promises";
const moduleModuleName = "module";

function resolveWasmPath(createRequire: (filename: string) => { resolve: (id: string) => string }) {
  // ESM: import.meta.url is available
  // CJS: tsup replaces import.meta with empty object, so fall back to __filename
  const baseUrl = import.meta.url || `file://${__filename}`;
  const require = createRequire(baseUrl);
  return require.resolve("@resvg/resvg-wasm/index_bg.wasm");
}

async function loadBundledWasm(): Promise<Uint8Array> {
  const [{ readFile }, { createRequire }] = await Promise.all([
    importNodeFsPromises(),
    importNodeModule(),
  ]);
  const wasmPath = resolveWasmPath(createRequire);
  return readFile(wasmPath);
}

function importNodeFsPromises(): Promise<typeof import("node:fs/promises")> {
  return import(`${nodeImportPrefix}${fsModuleName}/${fsPromisesSegment}`);
}

function importNodeModule(): Promise<typeof import("node:module")> {
  return import(`${nodeImportPrefix}${moduleModuleName}`);
}

async function normalizeWasmInput(wasm: ResvgWasmInput): Promise<ArrayBuffer | Uint8Array> {
  if (wasm instanceof ArrayBuffer || wasm instanceof Uint8Array) {
    return wasm;
  }
  return wasm.arrayBuffer();
}

/**
 * Initialize the resvg-wasm module used for PNG conversion.
 *
 * Calling this is optional because `convertPptxToPng` initializes resvg-wasm on
 * first use. Applications may call it during startup to pay the WASM loading
 * cost before handling the first conversion request. Pass the WASM binary
 * explicitly in browser-like environments where Node.js filesystem APIs are not
 * available.
 */
export async function initResvgWasm(wasm?: ResvgWasmInput): Promise<void> {
  if (!wasmInitPromise) {
    wasmInitPromise = (async () => {
      const wasmInput =
        wasm === undefined ? await loadBundledWasm() : await normalizeWasmInput(wasm);
      await initWasm(wasmInput);
    })();
  }
  await wasmInitPromise;
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<SvgToPngResult> {
  await initResvgWasm();
  return renderSvgToPng(svgString, options);
}

export type { PngConvertOptions, ResvgWasmInput, SvgToPngResult } from "./types.js";
