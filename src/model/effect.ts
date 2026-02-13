import type { ResolvedColor } from "./fill.js";

export interface EffectList {
  outerShadow: OuterShadow | null;
  innerShadow: InnerShadow | null;
  glow: Glow | null;
  softEdge: SoftEdge | null;
}

export interface OuterShadow {
  blurRadius: number;
  distance: number;
  direction: number;
  color: ResolvedColor;
  alignment: string;
  rotateWithShape: boolean;
}

export interface InnerShadow {
  blurRadius: number;
  distance: number;
  direction: number;
  color: ResolvedColor;
}

export interface Glow {
  radius: number;
  color: ResolvedColor;
}

export interface SoftEdge {
  radius: number;
}

export interface BlipEffects {
  grayscale: boolean;
  biLevel: BiLevelEffect | null;
  blur: BlurEffect | null;
  lum: LumEffect | null;
  duotone: DuotoneEffect | null;
}

export interface BiLevelEffect {
  threshold: number;
}

export interface BlurEffect {
  radius: number;
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
