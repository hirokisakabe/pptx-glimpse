import type { ShapeColorInput } from "./shape-xml.js";
import type { SourceRectangleAlignment } from "./shapes.js";
import type { Emu, OoxmlAngle } from "./units.js";

export interface OuterShadowInput {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: OoxmlAngle;
  readonly color: ShapeColorInput;
  readonly alignment: SourceRectangleAlignment;
  readonly rotateWithShape: boolean;
}

export interface InnerShadowInput {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: OoxmlAngle;
  readonly color: ShapeColorInput;
}

export interface ShadowEffectsInput {
  readonly outerShadow?: OuterShadowInput;
  readonly innerShadow?: InnerShadowInput;
}

const RECTANGLE_ALIGNMENTS: ReadonlySet<string> = new Set<SourceRectangleAlignment>([
  "tl",
  "t",
  "tr",
  "l",
  "ctr",
  "r",
  "bl",
  "b",
  "br",
]);
const MAX_POWERPOINT_POSITIVE_COORDINATE = 2147483647;

export function assertShadowEffectsInput(
  effects: unknown,
  operationName: "addShape" | "addPicture",
): void {
  if (!isPlainRecord(effects)) {
    throw new Error(`${operationName}: effects must be an object`);
  }
  if (effects.outerShadow !== undefined) {
    assertOuterShadowInput(effects.outerShadow, operationName);
  }
  if (effects.innerShadow !== undefined) {
    assertInnerShadowInput(effects.innerShadow, operationName);
  }
}

function assertOuterShadowInput(value: unknown, operationName: string): void {
  if (!isPlainRecord(value)) {
    throw new Error(`${operationName}: effects.outerShadow must be an object`);
  }
  assertShadowGeometry(value, operationName, "effects.outerShadow");
  if (typeof value.alignment !== "string" || !RECTANGLE_ALIGNMENTS.has(value.alignment)) {
    throw new Error(`${operationName}: effects.outerShadow.alignment is not supported`);
  }
  if (typeof value.rotateWithShape !== "boolean") {
    throw new Error(`${operationName}: effects.outerShadow.rotateWithShape must be a boolean`);
  }
}

function assertInnerShadowInput(value: unknown, operationName: string): void {
  if (!isPlainRecord(value)) {
    throw new Error(`${operationName}: effects.innerShadow must be an object`);
  }
  assertShadowGeometry(value, operationName, "effects.innerShadow");
}

function assertShadowGeometry(
  value: Record<string, unknown>,
  operationName: string,
  path: string,
): void {
  assertPowerPointPositiveCoordinate(value.blurRadius, operationName, `${path}.blurRadius`);
  assertPowerPointPositiveCoordinate(value.distance, operationName, `${path}.distance`);
  assertPositiveFixedAngle(value.direction, operationName, `${path}.direction`);
  assertShadowColor(value.color, operationName, `${path}.color`);
}

function assertShadowColor(value: unknown, operationName: string, path: string): void {
  if (!isPlainRecord(value) || value.kind !== "srgb") {
    throw new Error(`${operationName}: ${path} must be an srgb color object`);
  }
  if (typeof value.hex !== "string" || !/^[0-9A-Fa-f]{6}$/.test(value.hex)) {
    throw new Error(`${operationName}: ${path}.hex must be a 6-digit RGB value`);
  }
  if (value.transforms === undefined) return;
  if (!Array.isArray(value.transforms)) {
    throw new Error(`${operationName}: ${path}.transforms must be an array`);
  }
  value.transforms.forEach((transform, index) => {
    if (!isPlainRecord(transform) || transform.kind !== "alpha") {
      throw new Error(`${operationName}: ${path}.transforms[${index}].kind is not supported`);
    }
    if (
      typeof transform.value !== "number" ||
      !Number.isInteger(transform.value) ||
      transform.value < 0 ||
      transform.value > 100000
    ) {
      throw new Error(
        `${operationName}: ${path}.transforms[${index}].value must be between 0 and 100000`,
      );
    }
  });
}

function assertPowerPointPositiveCoordinate(
  value: unknown,
  operationName: string,
  path: string,
): void {
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value < 0 ||
    value > MAX_POWERPOINT_POSITIVE_COORDINATE
  ) {
    throw new Error(
      `${operationName}: ${path} must be an integer between 0 and ${MAX_POWERPOINT_POSITIVE_COORDINATE}`,
    );
  }
}

function assertPositiveFixedAngle(value: unknown, operationName: string, path: string): void {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value >= 21600000) {
    throw new Error(`${operationName}: ${path} must be an integer between 0 and 21599999`);
  }
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
