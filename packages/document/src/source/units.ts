import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";

/**
 * PptxSourceModel source model が使う PPTX-native な typed units。
 *
 * source model は pixel 変換を持たず、OOXML が定義する単位系をそのまま
 * 型として保持する。renderer 側の単位型 (`@pptx-glimpse/renderer`) は
 * 下位基盤である `@pptx-glimpse/document` から参照できないため、ここで
 * 独立に branded type を定義する。ランタイムコストはゼロ (JS 出力は
 * plain number と同一)。
 */

declare const EmuBrand: unique symbol;
declare const PtBrand: unique symbol;
declare const HundredthPtBrand: unique symbol;
declare const OoxmlPercentBrand: unique symbol;
declare const OoxmlAngleBrand: unique symbol;

/** English Metric Units (1 inch = 914,400 EMU)。座標・サイズに使う。 */
export type Emu = number & { readonly [EmuBrand]: typeof EmuBrand };

/** ポイント (1 pt = 1/72 inch)。フォントサイズに使う。 */
export type Pt = number & { readonly [PtBrand]: typeof PtBrand };

/** 1/100 ポイント (ECMA-376 `a:spcPts` 等)。 */
export type HundredthPt = number & { readonly [HundredthPtBrand]: typeof HundredthPtBrand };

/**
 * OOXML パーセンテージ (ST_Percentage)。整数値で 1% = 1000。
 * `a:spcPct` / `a:lumMod` / `a:tint` 等の transform で使う。
 */
export type OoxmlPercent = number & { readonly [OoxmlPercentBrand]: typeof OoxmlPercentBrand };

/**
 * OOXML 角度 (ST_Angle)。1 度 = 60,000。`a:xfrm@rot` 等で使う。
 */
export type OoxmlAngle = number & { readonly [OoxmlAngleBrand]: typeof OoxmlAngleBrand };

export function asEmu(value: number): Emu {
  return unsafeTypeAssertion<Emu>(value);
}

export function asPt(value: number): Pt {
  return unsafeTypeAssertion<Pt>(value);
}

export function asHundredthPt(value: number): HundredthPt {
  return unsafeTypeAssertion<HundredthPt>(value);
}

export function asOoxmlPercent(value: number): OoxmlPercent {
  return unsafeTypeAssertion<OoxmlPercent>(value);
}

export function asOoxmlAngle(value: number): OoxmlAngle {
  return unsafeTypeAssertion<OoxmlAngle>(value);
}
