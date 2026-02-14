import type { Fill } from "./fill.js";
import type { SlideElement } from "./shape.js";

export interface Slide {
  slideNumber: number;
  background: Background | null;
  elements: SlideElement[];
  showMasterSp: boolean;
}

export interface Background {
  fill: Fill | null;
}
