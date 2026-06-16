import { DEFAULT_DPI, EMU_PER_INCH, EMU_PER_POINT, ROTATION_UNIT } from "./constants.js";
import type { Emu, HundredthPt, Pt } from "./unit-types.js";

export function emuToPixels(emu: Emu, dpi: number = DEFAULT_DPI): number {
  return (emu / EMU_PER_INCH) * dpi;
}

export function emuToPoints(emu: Emu): number {
  return emu / EMU_PER_POINT;
}

/** 1/60000度 → 度 */
export function rotationToDegrees(rotation: number): number {
  return rotation / ROTATION_UNIT;
}

/** 1/100ポイント → ポイント */
export function hundredthPointToPoint(value: HundredthPt): Pt {
  return (value / 100) as Pt;
}
