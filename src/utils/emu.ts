import { DEFAULT_DPI, EMU_PER_INCH, EMU_PER_POINT, ROTATION_UNIT } from "./constants.js";

export function emuToPixels(emu: number, dpi: number = DEFAULT_DPI): number {
  return (emu / EMU_PER_INCH) * dpi;
}

export function emuToPoints(emu: number): number {
  return emu / EMU_PER_POINT;
}

/** 1/60000度 → 度 */
export function rotationToDegrees(rotation: number): number {
  return rotation / ROTATION_UNIT;
}

/** 1/100ポイント → ポイント */
export function hundredthPointToPoint(value: number): number {
  return value / 100;
}
