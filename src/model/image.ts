import type { Transform } from "./shape.js";

export interface ImageElement {
  type: "image";
  transform: Transform;
  imageData: string;
  mimeType: string;
}
