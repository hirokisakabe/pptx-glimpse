import { initWasm, type InitInput } from "@resvg/resvg-wasm";

let initialized = false;

export async function ensureWasmInitialized(): Promise<void> {
  if (initialized) return;

  // Node.js: auto-load WASM from the package directory
  if (typeof process !== "undefined" && process.versions?.node) {
    const { readFile } = await import("node:fs/promises");
    const { createRequire } = await import("node:module");
    // CJS: import.meta.url is empty, use __filename via createRequire workaround
    const req =
      typeof __filename !== "undefined"
        ? createRequire(__filename)
        : createRequire(import.meta.url);
    const wasmPath = req.resolve("@resvg/resvg-wasm/index_bg.wasm");
    const wasmBuffer = await readFile(wasmPath);
    await initWasm(wasmBuffer);
    initialized = true;
    return;
  }

  throw new Error(
    "WASM is not initialized. Call initPng() before using convertPptxToPng() in browser environments.",
  );
}

/**
 * Initialize the PNG conversion WASM module.
 *
 * - **Node.js**: Initialization is automatic. You do not need to call this function.
 * - **Browser**: You must call this function before using `convertPptxToPng()`.
 *
 * @example
 * ```ts
 * import { initPng, convertPptxToPng } from "pptx-glimpse";
 * await initPng(fetch("/path/to/index_bg.wasm"));
 * const result = await convertPptxToPng(uint8Array);
 * ```
 */
// eslint-disable-next-line @typescript-eslint/no-redundant-type-constituents
export async function initPng(wasmSource: Promise<InitInput> | InitInput): Promise<void> {
  if (initialized) return;
  await initWasm(wasmSource);
  initialized = true;
}
