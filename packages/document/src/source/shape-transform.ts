/**
 * Transform input shared by existing shape mutation and new-content shape authoring.
 */

import type { Emu } from "./units.js";

export interface UpdateShapeTransformInput {
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}
