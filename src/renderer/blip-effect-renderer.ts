import type { BlipEffects } from "../model/effect.js";
import { emuToPixels } from "../utils/emu.js";

interface BlipEffectRenderResult {
  filterAttr: string;
  filterDefs: string;
}

export function renderBlipEffects(blipEffects: BlipEffects | null): BlipEffectRenderResult {
  if (!blipEffects) return { filterAttr: "", filterDefs: "" };

  const primitives: string[] = [];
  let lastResult = "SourceGraphic";

  if (blipEffects.grayscale) {
    primitives.push(`<feColorMatrix type="saturate" values="0" result="grayscale"/>`);
    lastResult = "grayscale";
  }

  if (blipEffects.biLevel) {
    const thresh = blipEffects.biLevel.threshold;
    if (!blipEffects.grayscale) {
      primitives.push(
        `<feColorMatrix in="${lastResult}" type="saturate" values="0" result="biLevelGray"/>`,
      );
      lastResult = "biLevelGray";
    }
    const tv = buildThresholdTable(thresh);
    primitives.push(
      `<feComponentTransfer in="${lastResult}" result="biLevel">`,
      `<feFuncR type="discrete" tableValues="${tv}"/>`,
      `<feFuncG type="discrete" tableValues="${tv}"/>`,
      `<feFuncB type="discrete" tableValues="${tv}"/>`,
      `</feComponentTransfer>`,
    );
    lastResult = "biLevel";
  }

  if (blipEffects.blur) {
    const stdDev = emuToPixels(blipEffects.blur.radius) / 2;
    primitives.push(
      `<feGaussianBlur in="${lastResult}" stdDeviation="${stdDev}" result="blipBlur"/>`,
    );
    lastResult = "blipBlur";
  }

  if (blipEffects.lum) {
    const { brightness, contrast } = blipEffects.lum;
    const slope = round(1 + contrast);
    const intercept = round(brightness - contrast / 2);
    primitives.push(
      `<feComponentTransfer in="${lastResult}" result="lumResult">`,
      `<feFuncR type="linear" slope="${slope}" intercept="${intercept}"/>`,
      `<feFuncG type="linear" slope="${slope}" intercept="${intercept}"/>`,
      `<feFuncB type="linear" slope="${slope}" intercept="${intercept}"/>`,
      `</feComponentTransfer>`,
    );
    lastResult = "lumResult";
  }

  if (blipEffects.duotone) {
    const { color1, color2 } = blipEffects.duotone;
    const [r1, g1, b1] = hexToRgbNorm(color1.hex);
    const [r2, g2, b2] = hexToRgbNorm(color2.hex);
    if (!blipEffects.grayscale && !blipEffects.biLevel) {
      primitives.push(
        `<feColorMatrix in="${lastResult}" type="saturate" values="0" result="duotoneGray"/>`,
      );
      lastResult = "duotoneGray";
    }
    primitives.push(
      `<feComponentTransfer in="${lastResult}" result="duotoneResult">`,
      `<feFuncR type="table" tableValues="${r1} ${r2}"/>`,
      `<feFuncG type="table" tableValues="${g1} ${g2}"/>`,
      `<feFuncB type="table" tableValues="${b1} ${b2}"/>`,
      `</feComponentTransfer>`,
    );
    lastResult = "duotoneResult";
  }

  if (blipEffects.clrChange && blipEffects.clrChange.clrTo.alpha === 0) {
    // "Set Transparent Color" approximation using SVG feColorMatrix.
    // Computes alpha = scale * (R + G + B - Rf - Gf - Bf).
    // Pixels matching clrFrom (where R+G+B ≈ Rf+Gf+Bf) get alpha near 0.
    // This is an approximation: near-clrFrom colors become semi-transparent,
    // and colors darker than clrFrom may also lose alpha. Works well for the
    // common "white → transparent" case. Exact per-pixel color keying is not
    // achievable with SVG filters alone.
    // Only alpha=0 targets are supported; general color replacement is not.
    const { clrFrom } = blipEffects.clrChange;
    const [rf, gf, bf] = hexToRgbNorm(clrFrom.hex);
    const scale = 20;
    primitives.push(
      `<feColorMatrix in="${lastResult}" type="matrix" values="` +
        `1 0 0 0 0 ` +
        `0 1 0 0 0 ` +
        `0 0 1 0 0 ` +
        `${round(scale)} ${round(scale)} ${round(scale)} 0 ${round(-scale * (rf + gf + bf))}" result="clrChangeResult"/>`,
    );
    lastResult = "clrChangeResult";
  }

  if (primitives.length === 0) return { filterAttr: "", filterDefs: "" };

  const id = `blip-effect-${crypto.randomUUID()}`;
  const filterDefs = [
    `<filter id="${id}" color-interpolation-filters="sRGB">`,
    ...primitives,
    `</filter>`,
  ].join("");

  return {
    filterAttr: `filter="url(#${id})"`,
    filterDefs,
  };
}

function buildThresholdTable(threshold: number): string {
  const steps = 16;
  const values: number[] = [];
  for (let i = 0; i < steps; i++) {
    values.push(i / steps < threshold ? 0 : 1);
  }
  return values.join(" ");
}

function hexToRgbNorm(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [
    round(parseInt(h.substring(0, 2), 16) / 255),
    round(parseInt(h.substring(2, 4), 16) / 255),
    round(parseInt(h.substring(4, 6), 16) / 255),
  ];
}

function round(n: number): number {
  return Math.round(n * 1000) / 1000;
}
