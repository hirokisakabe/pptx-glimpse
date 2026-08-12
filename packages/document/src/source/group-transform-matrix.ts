import type { SourceTransform } from "./shapes.js";
import { asEmu, asOoxmlAngle } from "./units.js";

/** Source-local 2D affine matrix using the SVG/Canvas `(a, b, c, d, e, f)` layout. */
export interface SourceAffineMatrix {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

type SourceTransformMatrixRejection =
  | "missing-transform"
  | "missing-child-transform"
  | "non-finite-matrix"
  | "non-finite-transform"
  | "invalid-extent"
  | "zero-child-extent"
  | "singular-matrix"
  | "shear"
  | "out-of-range-transform"
  | "quantization-mismatch";

export type SourceTransformMatrixResult<T> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly reason: SourceTransformMatrixRejection };

const IDENTITY_MATRIX: SourceAffineMatrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
const OOXML_ANGLE_PER_RADIAN = (60000 * 180) / Math.PI;
const MATRIX_ABSOLUTE_TOLERANCE = 1e-5;
const MATRIX_ULP_TOLERANCE = 8;
const MATRIX_TOLERANCE_CAP = 0.01;
const NORMALIZED_SHEAR_TOLERANCE = 1e-10;
// ECMA-376 DrawingML ST_Coordinate / ST_PositiveCoordinate bounds. These are also safely
// representable as JavaScript integers, unlike the wider xsd:long storage type.
const OOXML_COORDINATE_MIN = -27_273_042_329_600;
const OOXML_COORDINATE_MAX = 27_273_042_316_900;
const OOXML_POSITIVE_COORDINATE_MAX = 27_273_042_316_900;

export interface SourceGroupCoordinateSpace {
  readonly transform?: SourceTransform;
  readonly childTransform?: SourceTransform;
}

/** Multiplies matrices so the returned mapping applies `inner` before `outer`. */
export function multiplySourceAffineMatrices(
  outer: SourceAffineMatrix,
  inner: SourceAffineMatrix,
): SourceAffineMatrix {
  return {
    a: outer.a * inner.a + outer.c * inner.b,
    b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d,
    d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e,
    f: outer.b * inner.e + outer.d * inner.f + outer.f,
  };
}

/** Composes a source-to-child ancestor chain in authored outer-to-inner order. */
export function composeSourceAncestorMatrices(
  outerToInner: readonly SourceAffineMatrix[],
): SourceTransformMatrixResult<SourceAffineMatrix> {
  if (!outerToInner.every(isFiniteMatrix)) return reject("non-finite-matrix");
  const composed = outerToInner.reduce(
    (current, inner) => multiplySourceAffineMatrices(current, inner),
    IDENTITY_MATRIX,
  );
  return isFiniteMatrix(composed) ? accept(composed) : reject("non-finite-matrix");
}

/**
 * Builds the immediate-child to parent mapping authored by a native group:
 * `T(off) * R(ext center) * F(ext center) * S(ext/chExt) * T(-chOff)`.
 */
export function buildSourceGroupMappingMatrix(
  transform: SourceTransform | undefined,
  childTransform: SourceTransform | undefined,
): SourceTransformMatrixResult<SourceAffineMatrix> {
  if (transform === undefined) return reject("missing-transform");
  if (childTransform === undefined) return reject("missing-child-transform");
  if (!isFiniteTransform(transform) || !isFiniteTransform(childTransform)) {
    return reject("non-finite-transform");
  }
  if (Number(transform.width) < 0 || Number(transform.height) < 0) {
    return reject("invalid-extent");
  }

  const childWidth = Number(childTransform.width);
  const childHeight = Number(childTransform.height);
  if (childWidth === 0 || childHeight === 0) return reject("zero-child-extent");
  if (childWidth < 0 || childHeight < 0) return reject("invalid-extent");

  const width = Number(transform.width);
  const height = Number(transform.height);
  const offsetX = Number(transform.offsetX);
  const offsetY = Number(transform.offsetY);
  const centerX = width / 2;
  const centerY = height / 2;
  const radians = ooxmlAngleToRadians(Number(transform.rotation ?? 0));
  const flipX = transform.flipHorizontal === true ? -1 : 1;
  const flipY = transform.flipVertical === true ? -1 : 1;

  const matrix = multiplySourceAffineMatrices(
    translationMatrix(offsetX, offsetY),
    multiplySourceAffineMatrices(
      rotationAroundPointMatrix(radians, centerX, centerY),
      multiplySourceAffineMatrices(
        flipAroundExtentCenterMatrix(flipX, flipY, width, height),
        multiplySourceAffineMatrices(
          scaleMatrix(width / childWidth, height / childHeight),
          translationMatrix(-Number(childTransform.offsetX), -Number(childTransform.offsetY)),
        ),
      ),
    ),
  );
  return isFiniteMatrix(matrix) ? accept(matrix) : reject("non-finite-matrix");
}

export function invertSourceAffineMatrix(
  matrix: SourceAffineMatrix,
): SourceTransformMatrixResult<SourceAffineMatrix> {
  if (!isFiniteMatrix(matrix)) return reject("non-finite-matrix");
  const linearScale = Math.max(
    Math.abs(matrix.a),
    Math.abs(matrix.b),
    Math.abs(matrix.c),
    Math.abs(matrix.d),
  );
  if (linearScale === 0) return reject("singular-matrix");
  const normalizedA = matrix.a / linearScale;
  const normalizedB = matrix.b / linearScale;
  const normalizedC = matrix.c / linearScale;
  const normalizedD = matrix.d / linearScale;
  const normalizedDeterminant = normalizedA * normalizedD - normalizedB * normalizedC;
  if (Math.abs(normalizedDeterminant) <= Number.EPSILON * 8) {
    return reject("singular-matrix");
  }

  const inverseScale = 1 / linearScale / normalizedDeterminant;
  const inverse = {
    a: normalizedD * inverseScale,
    b: -normalizedB * inverseScale,
    c: -normalizedC * inverseScale,
    d: normalizedA * inverseScale,
    e: 0,
    f: 0,
  };
  inverse.e = -(inverse.a * matrix.e + inverse.c * matrix.f);
  inverse.f = -(inverse.b * matrix.e + inverse.d * matrix.f);
  return isFiniteMatrix(inverse) ? accept(inverse) : reject("non-finite-matrix");
}

/** Builds a normalized-unit-rectangle to parent-local mapping for an OOXML transform. */
export function buildSourceTransformMatrix(
  transform: SourceTransform,
): SourceTransformMatrixResult<SourceAffineMatrix> {
  if (!isFiniteTransform(transform)) return reject("non-finite-transform");
  const width = Number(transform.width);
  const height = Number(transform.height);
  if (width <= 0 || height <= 0) return reject("invalid-extent");

  const radians = ooxmlAngleToRadians(Number(transform.rotation ?? 0));
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  const flipX = transform.flipHorizontal === true ? -1 : 1;
  const flipY = transform.flipVertical === true ? -1 : 1;
  const a = cosine * flipX * width;
  const b = sine * flipX * width;
  const c = -sine * flipY * height;
  const d = cosine * flipY * height;
  const centerX = width / 2;
  const centerY = height / 2;
  const matrix = {
    a,
    b,
    c,
    d,
    e: Number(transform.offsetX) + centerX - (a + c) / 2,
    f: Number(transform.offsetY) + centerY - (b + d) / 2,
  };
  return isFiniteMatrix(matrix) ? accept(matrix) : reject("non-finite-matrix");
}

/**
 * Re-expresses one drawing root transform from a source parent coordinate space into a
 * destination parent coordinate space without changing its part-absolute mapping.
 */
export function reparentSourceTransform(
  sourceOuterToInner: readonly SourceGroupCoordinateSpace[],
  destinationOuterToInner: readonly SourceGroupCoordinateSpace[],
  transform: SourceTransform | undefined,
): SourceTransformMatrixResult<SourceTransform> {
  if (transform === undefined) return reject("missing-transform");
  const sourceParent = composeGroupCoordinateSpaces(sourceOuterToInner);
  if (!sourceParent.ok) return sourceParent;
  const destinationParent = composeGroupCoordinateSpaces(destinationOuterToInner);
  if (!destinationParent.ok) return destinationParent;
  const destinationInverse = invertSourceAffineMatrix(destinationParent.value);
  if (!destinationInverse.ok) return destinationInverse;
  const local = buildSourceTransformMatrix(transform);
  if (!local.ok) return local;
  const reparented = multiplySourceAffineMatrices(
    destinationInverse.value,
    multiplySourceAffineMatrices(sourceParent.value, local.value),
  );
  return decomposeSourceTransformMatrix(reparented);
}

function composeGroupCoordinateSpaces(
  groups: readonly SourceGroupCoordinateSpace[],
): SourceTransformMatrixResult<SourceAffineMatrix> {
  const matrices: SourceAffineMatrix[] = [];
  for (const group of groups) {
    const matrix = buildSourceGroupMappingMatrix(group.transform, group.childTransform);
    if (!matrix.ok) return matrix;
    matrices.push(matrix.value);
  }
  return composeSourceAncestorMatrices(matrices);
}

/**
 * Returns an integer-EMU/integer-angle OOXML transform only when rebuilding that
 * quantized transform reproduces the input matrix within floating-point tolerance.
 * Reflections use a canonical horizontal flip; a vertical reflection remains exact by
 * combining that flip with an equivalent rotation.
 */
export function decomposeSourceTransformMatrix(
  matrix: SourceAffineMatrix,
): SourceTransformMatrixResult<SourceTransform> {
  if (!isFiniteMatrix(matrix)) return reject("non-finite-matrix");

  const width = Math.hypot(matrix.a, matrix.b);
  const height = Math.hypot(matrix.c, matrix.d);
  if (width === 0 || height === 0) return reject("singular-matrix");

  const normalizedDot = (matrix.a * matrix.c + matrix.b * matrix.d) / (width * height);
  if (Math.abs(normalizedDot) > NORMALIZED_SHEAR_TOLERANCE) return reject("shear");

  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  if (!Number.isFinite(determinant) || determinant === 0) return reject("singular-matrix");
  const flipHorizontal = determinant < 0;
  const rotationRadians = flipHorizontal
    ? Math.atan2(-matrix.b, -matrix.a)
    : Math.atan2(matrix.b, matrix.a);
  const offsetX = matrix.e - width / 2 + (matrix.a + matrix.c) / 2;
  const offsetY = matrix.f - height / 2 + (matrix.b + matrix.d) / 2;

  const quantized: SourceTransform = {
    offsetX: asEmu(Math.round(offsetX)),
    offsetY: asEmu(Math.round(offsetY)),
    width: asEmu(Math.round(width)),
    height: asEmu(Math.round(height)),
    rotation: asOoxmlAngle(Math.round(rotationRadians * OOXML_ANGLE_PER_RADIAN)),
    ...(flipHorizontal ? { flipHorizontal: true } : {}),
  };
  if (!isQuantizedTransformInOoxmlRange(quantized)) {
    return reject("out-of-range-transform");
  }
  const reconstructed = buildSourceTransformMatrix(quantized);
  if (!reconstructed.ok) return reconstructed;
  return matricesNearlyEqual(matrix, reconstructed.value)
    ? accept(quantized)
    : reject("quantization-mismatch");
}

function isQuantizedTransformInOoxmlRange(transform: SourceTransform): boolean {
  const offsetX = Number(transform.offsetX);
  const offsetY = Number(transform.offsetY);
  const width = Number(transform.width);
  const height = Number(transform.height);
  const rotation = Number(transform.rotation ?? 0);
  return (
    [offsetX, offsetY, width, height, rotation].every(Number.isSafeInteger) &&
    offsetX >= OOXML_COORDINATE_MIN &&
    offsetX <= OOXML_COORDINATE_MAX &&
    offsetY >= OOXML_COORDINATE_MIN &&
    offsetY <= OOXML_COORDINATE_MAX &&
    width > 0 &&
    width <= OOXML_POSITIVE_COORDINATE_MAX &&
    height > 0 &&
    height <= OOXML_POSITIVE_COORDINATE_MAX
  );
}

function matricesNearlyEqual(left: SourceAffineMatrix, right: SourceAffineMatrix): boolean {
  return (["a", "b", "c", "d", "e", "f"] as const).every((key) => {
    const floatingPointTolerance = Math.min(
      MATRIX_TOLERANCE_CAP,
      MATRIX_ULP_TOLERANCE * Math.max(floatingPointUlp(left[key]), floatingPointUlp(right[key])),
    );
    const tolerance = Math.max(MATRIX_ABSOLUTE_TOLERANCE, floatingPointTolerance);
    return Math.abs(left[key] - right[key]) <= tolerance;
  });
}

function floatingPointUlp(value: number): number {
  const absolute = Math.abs(value);
  if (absolute === 0) return Number.MIN_VALUE;
  if (absolute < 2 ** -1022) return Number.MIN_VALUE;
  return 2 ** (Math.floor(Math.log2(absolute)) - 52);
}

function isFiniteTransform(transform: SourceTransform): boolean {
  return [
    transform.offsetX,
    transform.offsetY,
    transform.width,
    transform.height,
    transform.rotation ?? 0,
  ].every((value) => Number.isFinite(Number(value)));
}

function isFiniteMatrix(matrix: SourceAffineMatrix): boolean {
  return [matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f].every(Number.isFinite);
}

function translationMatrix(x: number, y: number): SourceAffineMatrix {
  return { a: 1, b: 0, c: 0, d: 1, e: x, f: y };
}

function scaleMatrix(x: number, y: number): SourceAffineMatrix {
  return { a: x, b: 0, c: 0, d: y, e: 0, f: 0 };
}

function rotationAroundPointMatrix(
  radians: number,
  centerX: number,
  centerY: number,
): SourceAffineMatrix {
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  return {
    a: cosine,
    b: sine,
    c: -sine,
    d: cosine,
    e: centerX - cosine * centerX + sine * centerY,
    f: centerY - sine * centerX - cosine * centerY,
  };
}

function flipAroundExtentCenterMatrix(
  flipX: number,
  flipY: number,
  width: number,
  height: number,
): SourceAffineMatrix {
  return {
    a: flipX,
    b: 0,
    c: 0,
    d: flipY,
    e: flipX === -1 ? width : 0,
    f: flipY === -1 ? height : 0,
  };
}

function ooxmlAngleToRadians(angle: number): number {
  return angle / OOXML_ANGLE_PER_RADIAN;
}

function accept<T>(value: T): SourceTransformMatrixResult<T> {
  return { ok: true, value };
}

function reject<T>(reason: SourceTransformMatrixRejection): SourceTransformMatrixResult<T> {
  return { ok: false, reason };
}
