import type { Geometry } from "../../model/shape.js";
import { getPresetGeometrySvg } from "./preset-geometries.js";

export function renderGeometry(geometry: Geometry, width: number, height: number): string {
  if (geometry.type === "preset") {
    return getPresetGeometrySvg(geometry.preset, width, height, geometry.adjustValues);
  }
  if (geometry.type === "custom" && geometry.pathData) {
    return `<path d="${geometry.pathData}"/>`;
  }
  return `<rect width="${width}" height="${height}"/>`;
}
