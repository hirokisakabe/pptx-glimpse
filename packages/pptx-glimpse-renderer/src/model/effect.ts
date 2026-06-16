import type { Emu } from "../utils/unit-types.js";
import type { ResolvedColor } from "./fill.js";

export interface EffectList {
  outerShadow: OuterShadow | null;
  innerShadow: InnerShadow | null;
  glow: Glow | null;
  softEdge: SoftEdge | null;
}

export interface OuterShadow {
  blurRadius: Emu;
  distance: Emu;
  direction: number;
  color: ResolvedColor;
  alignment: string;
  rotateWithShape: boolean;
}

export interface InnerShadow {
  blurRadius: Emu;
  distance: Emu;
  direction: number;
  color: ResolvedColor;
}

export interface Glow {
  radius: Emu;
  color: ResolvedColor;
}

export interface SoftEdge {
  radius: Emu;
}

export interface BlipEffects {
  grayscale: boolean;
  biLevel: BiLevelEffect | null;
  blur: BlurEffect | null;
  lum: LumEffect | null;
  duotone: DuotoneEffect | null;
  clrChange: ClrChangeEffect | null;
}

export interface BiLevelEffect {
  threshold: number;
}

export interface BlurEffect {
  radius: Emu;
  grow: boolean;
}

export interface LumEffect {
  brightness: number;
  contrast: number;
}

export interface DuotoneEffect {
  color1: ResolvedColor;
  color2: ResolvedColor;
}

export interface ClrChangeEffect {
  clrFrom: ResolvedColor;
  clrTo: ResolvedColor;
}
