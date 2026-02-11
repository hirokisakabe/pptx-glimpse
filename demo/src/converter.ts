import { convertPptxToSvg } from "../../src/converter.js";
import type { SlideSvg } from "../../src/converter.js";

export type { SlideSvg };

export async function convertFile(file: File): Promise<SlideSvg[]> {
  const arrayBuffer = await file.arrayBuffer();
  const uint8Array = new Uint8Array(arrayBuffer);
  return convertPptxToSvg(uint8Array);
}
