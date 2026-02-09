import type { Geometry } from "../../model/shape.js";
import { getPresetGeometrySvg } from "./preset-geometries.js";

export function renderGeometry(geometry: Geometry, width: number, height: number): string {
  if (geometry.type === "preset") {
    return getPresetGeometrySvg(geometry.preset, width, height, geometry.adjustValues);
  }
  // custom geometry fallback
  return `<rect width="${width}" height="${height}"/>`;
}
