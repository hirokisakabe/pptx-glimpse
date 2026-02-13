import type { Transform } from "./shape.js";
import type { EffectList, BlipEffects } from "./effect.js";

export interface SrcRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface StretchFillRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface TileInfo {
  tx: number;
  ty: number;
  sx: number;
  sy: number;
  flip: "none" | "x" | "y" | "xy";
  align: string;
}

export interface ImageElement {
  type: "image";
  transform: Transform;
  imageData: string;
  mimeType: string;
  effects: EffectList | null;
  blipEffects: BlipEffects | null;
  srcRect: SrcRect | null;
  stretch: StretchFillRect | null;
  tile: TileInfo | null;
}
