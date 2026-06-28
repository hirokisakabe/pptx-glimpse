import { unsafeBrandAssertion } from "../unsafe-type-assertion.js";

/**
 * Internal note.
 *
 * The source model does not perform pixel conversion and keeps the units defined by OOXML
 * Internal note.
 * the lower-level foundation `@pptx-glimpse/document` cannot reference, so this module
 * Internal note.
 * Internal note.
 */

declare const EmuBrand: unique symbol;
declare const PtBrand: unique symbol;
declare const HundredthPtBrand: unique symbol;
declare const OoxmlPercentBrand: unique symbol;
declare const OoxmlAngleBrand: unique symbol;

/** Internal note. */
export type Emu = number & { readonly [EmuBrand]: typeof EmuBrand };

/** points (1 pt = 1/72 inch)。font size. */
export type Pt = number & { readonly [PtBrand]: typeof PtBrand };

/** Internal note. */
export type HundredthPt = number & { readonly [HundredthPtBrand]: typeof HundredthPtBrand };

/**
 * OOXML percentage (ST_Percentage)。Integer value where 1% = 1000.
 * `a:spcPct` / `a:lumMod` / `a:tint` and similar transforms.
 */
export type OoxmlPercent = number & { readonly [OoxmlPercentBrand]: typeof OoxmlPercentBrand };

/**
 * OOXML angle (ST_Angle)。1 degree = 60,000.`a:xfrm@rot` and similar fields.
 */
export type OoxmlAngle = number & { readonly [OoxmlAngleBrand]: typeof OoxmlAngleBrand };

export function asEmu(value: number): Emu {
  return unsafeBrandAssertion<Emu>(value);
}

export function asPt(value: number): Pt {
  return unsafeBrandAssertion<Pt>(value);
}

export function asHundredthPt(value: number): HundredthPt {
  return unsafeBrandAssertion<HundredthPt>(value);
}

export function asOoxmlPercent(value: number): OoxmlPercent {
  return unsafeBrandAssertion<OoxmlPercent>(value);
}

export function asOoxmlAngle(value: number): OoxmlAngle {
  return unsafeBrandAssertion<OoxmlAngle>(value);
}
