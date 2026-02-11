import type { Transform } from "./shape.js";
import type { EffectList } from "./effect.js";

export interface ImageElement {
  type: "image";
  transform: Transform;
  imageData: string;
  mimeType: string;
  effects: EffectList | null;
}
