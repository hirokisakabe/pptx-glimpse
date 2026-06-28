import { unsafeBrandAssertion } from "../unsafe-type-assertion.js";

/**
 * Internal note.
 *
 * Internal note.
 * Internal note.
 * Internal note.
 */

declare const EmuBrand: unique symbol;
declare const PtBrand: unique symbol;
declare const HundredthPtBrand: unique symbol;

/** English Metric Units (1 inch = 914,400 EMU) */
export type Emu = number & { readonly [EmuBrand]: typeof EmuBrand };

/** points (1 pt = 1/72 inch) */
export type Pt = number & { readonly [PtBrand]: typeof PtBrand };

/** Internal note. */
export type HundredthPt = number & { readonly [HundredthPtBrand]: typeof HundredthPtBrand };

export function asEmu(value: number): Emu {
  return unsafeBrandAssertion<Emu>(value);
}

export function asHundredthPt(value: number): HundredthPt {
  return unsafeBrandAssertion<HundredthPt>(value);
}
