import type { Slide } from "./slide.js";
import type { Theme, ColorMap } from "./theme.js";
import type { DefaultTextStyle } from "./text.js";

export interface Presentation {
  slideSize: SlideSize;
  slides: Slide[];
  theme: Theme;
  colorMap: ColorMap;
  defaultTextStyle?: DefaultTextStyle;
}

export interface SlideSize {
  width: number;
  height: number;
}
