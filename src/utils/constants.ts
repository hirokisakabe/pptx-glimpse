// OOXML 単位系 (ECMA-376 §20.1.2.1)
// EMU (English Metric Units) は OOXML 内部の座標単位。
// 16:9 スライド = 9,144,000 × 5,143,500 EMU = 960 × 540 px (96 DPI)

/** 1 inch = 914,400 EMU */
export const EMU_PER_INCH = 914400;
/** 1 pt = 12,700 EMU */
export const EMU_PER_POINT = 12700;
export const DEFAULT_DPI = 96;
export const DEFAULT_OUTPUT_WIDTH = 960;

/** 回転角度の単位: 1/60,000 度 (ECMA-376 §20.1.10.3) */
export const ROTATION_UNIT = 60000;

/** フォントサイズの単位: 1/100 ポイント (ECMA-376 §21.1.2.2.25 sz 属性) */
export const FONT_SIZE_UNIT = 100;
