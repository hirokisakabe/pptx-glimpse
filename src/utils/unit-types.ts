/**
 * ブランド型（Branded Types）による単位の型安全性
 *
 * PPTX 内部で使われる複数の単位系（EMU, pt, 1/100 pt）を
 * コンパイル時に区別し、単位の取り違えを防止する。
 * ランタイムコストはゼロ（JS 出力は plain number と同一）。
 */

declare const EmuBrand: unique symbol;
declare const PtBrand: unique symbol;
declare const HundredthPtBrand: unique symbol;

/** English Metric Units (1 inch = 914,400 EMU) */
export type Emu = number & { readonly [EmuBrand]: typeof EmuBrand };

/** ポイント (1 pt = 1/72 inch) */
export type Pt = number & { readonly [PtBrand]: typeof PtBrand };

/** 1/100 ポイント (ECMA-376 spcPts 等) */
export type HundredthPt = number & { readonly [HundredthPtBrand]: typeof HundredthPtBrand };

export function asEmu(value: number): Emu {
  return value as Emu;
}

export function asPt(value: number): Pt {
  return value as Pt;
}

export function asHundredthPt(value: number): HundredthPt {
  return value as HundredthPt;
}
