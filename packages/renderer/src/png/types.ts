export type ResvgWasmInput = ArrayBuffer | Uint8Array | Response;

export interface PngConvertOptions {
  width?: number;
  height?: number;
  /** Font buffers used to render SVG <text> elements */
  fontBuffers?: Uint8Array[];
}

export interface SvgToPngResult {
  png: Uint8Array;
  width: number;
  height: number;
}
