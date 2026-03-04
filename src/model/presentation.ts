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

export interface SlideSize {
  width: Emu;
  height: Emu;
}
