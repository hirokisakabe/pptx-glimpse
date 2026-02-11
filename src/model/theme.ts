export interface Theme {
  colorScheme: ColorScheme;
  fontScheme: FontScheme;
}

export interface ColorScheme {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
}

export type ColorSchemeKey = keyof ColorScheme;

export interface ColorMap {
  bg1: ColorSchemeKey;
  tx1: ColorSchemeKey;
  bg2: ColorSchemeKey;
  tx2: ColorSchemeKey;
  accent1: ColorSchemeKey;
  accent2: ColorSchemeKey;
  accent3: ColorSchemeKey;
  accent4: ColorSchemeKey;
  accent5: ColorSchemeKey;
  accent6: ColorSchemeKey;
  hlink: ColorSchemeKey;
  folHlink: ColorSchemeKey;
}

export interface FontScheme {
  majorFont: string;
  minorFont: string;
  majorFontEa: string | null;
  minorFontEa: string | null;
}
