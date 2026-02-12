import type { EffectList } from "../model/effect.js";
import { emuToPixels } from "../utils/emu.js";

export interface EffectRenderResult {
  filterAttr: string;
  filterDefs: string;
}

export function renderEffects(effects: EffectList | null): EffectRenderResult {
  if (!effects) {
    return { filterAttr: "", filterDefs: "" };
  }

  const primitives: string[] = [];
  let lastResult = "SourceGraphic";

  if (effects.softEdge) {
    const r = emuToPixels(effects.softEdge.radius);
    primitives.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${r}" result="softEdgeMask"/>`,
      `<feComposite in="SourceGraphic" in2="softEdgeMask" operator="in" result="softEdgeResult"/>`,
    );
    lastResult = "softEdgeResult";
  }

  if (effects.glow) {
    const r = emuToPixels(effects.glow.radius);
    const { hex, alpha } = effects.glow.color;
    if (lastResult !== "SourceGraphic") {
      primitives.push(
        `<feColorMatrix in="${lastResult}" type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 1 0" result="glowAlpha"/>`,
      );
    }
    const blurIn = lastResult === "SourceGraphic" ? "SourceAlpha" : "glowAlpha";
    const mergeIn = lastResult;
    primitives.push(
      `<feGaussianBlur in="${blurIn}" stdDeviation="${r}" result="glowBlur"/>`,
      `<feFlood flood-color="${hex}" flood-opacity="${alpha}" result="glowColor"/>`,
      `<feComposite in="glowColor" in2="glowBlur" operator="in" result="glowFinal"/>`,
      `<feMerge result="glowMerge">`,
      `<feMergeNode in="glowFinal"/>`,
      `<feMergeNode in="${mergeIn}"/>`,
      `</feMerge>`,
    );
    lastResult = "glowMerge";
  }

  if (effects.outerShadow) {
    const stdDev = emuToPixels(effects.outerShadow.blurRadius) / 2;
    const dist = emuToPixels(effects.outerShadow.distance);
    const dirRad = (effects.outerShadow.direction * Math.PI) / 180;
    const dx = round(dist * Math.cos(dirRad));
    const dy = round(dist * Math.sin(dirRad));
    const { hex, alpha } = effects.outerShadow.color;
    const mergeIn = lastResult;

    primitives.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${stdDev}" result="shadowBlur"/>`,
      `<feOffset in="shadowBlur" dx="${dx}" dy="${dy}" result="shadowOffset"/>`,
      `<feFlood flood-color="${hex}" flood-opacity="${alpha}" result="shadowColor"/>`,
      `<feComposite in="shadowColor" in2="shadowOffset" operator="in" result="shadowFinal"/>`,
      `<feMerge result="outerShadowMerge">`,
      `<feMergeNode in="shadowFinal"/>`,
      `<feMergeNode in="${mergeIn}"/>`,
      `</feMerge>`,
    );
    lastResult = "outerShadowMerge";
  }

  if (effects.innerShadow) {
    const stdDev = emuToPixels(effects.innerShadow.blurRadius) / 2;
    const dist = emuToPixels(effects.innerShadow.distance);
    const dirRad = (effects.innerShadow.direction * Math.PI) / 180;
    const dx = round(dist * Math.cos(dirRad));
    const dy = round(dist * Math.sin(dirRad));
    const { hex, alpha } = effects.innerShadow.color;
    const sourceIn = lastResult;

    primitives.push(
      `<feComponentTransfer in="SourceAlpha" result="innerShdwInverse">`,
      `<feFuncA type="table" tableValues="1 0"/>`,
      `</feComponentTransfer>`,
      `<feGaussianBlur in="innerShdwInverse" stdDeviation="${stdDev}" result="innerShdwBlur"/>`,
      `<feOffset in="innerShdwBlur" dx="${dx}" dy="${dy}" result="innerShdwOffset"/>`,
      `<feFlood flood-color="${hex}" flood-opacity="${alpha}" result="innerShdwFill"/>`,
      `<feComposite in="innerShdwFill" in2="innerShdwOffset" operator="in" result="innerShdwColored"/>`,
      `<feComposite in="innerShdwColored" in2="SourceAlpha" operator="in" result="innerShdwClipped"/>`,
      `<feComposite in="innerShdwClipped" in2="${sourceIn}" operator="over"/>`,
    );
    lastResult = "";
  }

  if (primitives.length === 0) {
    return { filterAttr: "", filterDefs: "" };
  }

  const id = `effect-${crypto.randomUUID()}`;
  const filterDefs = [
    `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%" color-interpolation-filters="sRGB">`,
    ...primitives,
    `</filter>`,
  ].join("");

  return {
    filterAttr: `filter="url(#${id})"`,
    filterDefs,
  };
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}
