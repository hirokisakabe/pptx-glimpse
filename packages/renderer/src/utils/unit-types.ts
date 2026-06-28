import { unsafeBrandAssertion } from "../unsafe-type-assertion.js";

/**
 * Type safety of units with branded types
 *
 * Multiple unit systems used internally in PPTX (EMU, pt, 1/100 pt)
 * Distinguish at compile time to prevent unit mix-ups.
 * Runtime cost is zero (JS output is the same as plain number).
 */

declare const EmuBrand: unique symbol;
declare const PtBrand: unique symbol;
declare const HundredthPtBrand: unique symbol;

/** English Metric Units (1 inch = 914,400 EMU) */
export type Emu = number & { readonly [EmuBrand]: typeof EmuBrand };

/** points (1 pt = 1/72 inch) */
export type Pt = number & { readonly [PtBrand]: typeof PtBrand };

/** 1/100 point (ECMA-376 spcPts etc.) */
export type HundredthPt = number & { readonly [HundredthPtBrand]: typeof HundredthPtBrand };

export function asEmu(value: number): Emu {
  return unsafeBrandAssertion<Emu>(value);
}

export function asHundredthPt(value: number): HundredthPt {
  return unsafeBrandAssertion<HundredthPt>(value);
}
