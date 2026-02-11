import type { Geometry, CustomGeometryPath } from "../../model/shape.js";
import { getPresetGeometrySvg } from "./preset-geometries.js";

export function renderGeometry(geometry: Geometry, width: number, height: number): string {
  if (geometry.type === "preset") {
    return getPresetGeometrySvg(geometry.preset, width, height, geometry.adjustValues);
  }
  if (geometry.type === "custom" && geometry.paths.length > 0) {
    return renderCustomGeometry(geometry.paths, width, height);
  }
  return `<rect width="${width}" height="${height}"/>`;
}

function renderCustomGeometry(
  paths: CustomGeometryPath[],
  shapeWidth: number,
  shapeHeight: number,
): string {
  if (paths.length === 1) {
    return renderCustomPath(paths[0], shapeWidth, shapeHeight);
  }

  const parts: string[] = ["<g>"];
  for (const path of paths) {
    parts.push(renderCustomPath(path, shapeWidth, shapeHeight));
  }
  parts.push("</g>");
  return parts.join("");
}

function renderCustomPath(
  path: CustomGeometryPath,
  shapeWidth: number,
  shapeHeight: number,
): string {
  const scaleX = path.width > 0 ? shapeWidth / path.width : 1;
  const scaleY = path.height > 0 ? shapeHeight / path.height : 1;
  return `<path d="${path.commands}" transform="scale(${scaleX}, ${scaleY})"/>`;
}
