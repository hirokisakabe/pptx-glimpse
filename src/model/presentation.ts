import type { Slide } from "./slide.js";
import type { Theme, ColorMap } from "./theme.js";
import type { DefaultTextStyle, TxStyles } from "./text.js";
import type { Emu } from "../utils/unit-types.js";

export interface EmbeddedFont {
  typeface: string;
  panose?: string;
  pitchFamily?: number;
  charset?: number;
}

export interface Protection {
  modifyVerifier?: {
    algorithmName?: string;
    hashValue?: string;
    saltValue?: string;
    spinCount?: number;
  };
}

export interface Presentation {
  slideSize: SlideSize;
  slides: Slide[];
  theme: Theme;
  colorMap: ColorMap;
  defaultTextStyle?: DefaultTextStyle;
  txStyles?: TxStyles;
  embeddedFonts?: EmbeddedFont[];
  protection?: Protection;
}

export interface SlideSize {
  width: Emu;
  height: Emu;
}
