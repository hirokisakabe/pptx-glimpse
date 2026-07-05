import { initWasm } from "@resvg/resvg-wasm";

import { renderSvgToPng } from "./render-svg.js";
import type { PngConvertOptions, ResvgWasmInput, SvgToPngResult } from "./types.js";

let wasmInitPromise: Promise<void> | null = null;

async function normalizeWasmInput(wasm: ResvgWasmInput): Promise<ArrayBuffer | Uint8Array> {
  if (wasm instanceof ArrayBuffer || wasm instanceof Uint8Array) {
    return wasm;
  }
  if (!wasm.ok) {
    throw new Error(`Failed to load resvg WASM: HTTP ${wasm.status.toString()}`);
  }
  return wasm.arrayBuffer();
}

export async function initResvgWasm(wasm: ResvgWasmInput): Promise<void> {
  if (!wasmInitPromise) {
    wasmInitPromise = normalizeWasmInput(wasm)
      .then((wasmInput) => initWasm(wasmInput))
      .catch((error: unknown) => {
        wasmInitPromise = null;
        throw error;
      });
  }
  await wasmInitPromise;
}

export async function svgToPng(
  svgString: string,
  options?: PngConvertOptions,
): Promise<SvgToPngResult> {
  if (!wasmInitPromise) {
    throw new Error("initResvgWasm(wasm) must be called before browser PNG conversion.");
  }
  await wasmInitPromise;
  return renderSvgToPng(svgString, options);
}

export type { PngConvertOptions, ResvgWasmInput, SvgToPngResult } from "./types.js";
