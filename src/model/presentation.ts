import type { Slide } from "./slide.js";
import type { Theme, ColorMap } from "./theme.js";

export interface Presentation {
  slideSize: SlideSize;
  slides: Slide[];
  theme: Theme;
  colorMap: ColorMap;
}

export interface SlideSize {
  width: number;
  height: number;
}
