import type { Transform } from "./shape.js";
import type { EffectList } from "./effect.js";

export interface SrcRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface ImageElement {
  type: "image";
  transform: Transform;
  imageData: string;
  mimeType: string;
  effects: EffectList | null;
  srcRect: SrcRect | null;
}
