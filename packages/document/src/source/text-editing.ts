import { sourceHandlesEqual } from "./edit-descriptors.js";
import { asSourceNodeId } from "./handles.js";
import type {
  EditableParagraphProperties,
  EditableParagraphProperty,
  EditableTextRunProperties,
  EditableTextRunProperty,
  PptxSourceModel,
  SourceAutoNumScheme,
  SourceHandle,
  SourceParagraph,
  SourceParagraphProperties,
  SourceRunProperties,
  SourceShape,
  SourceSlide,
  SourceTextAlign,
  SourceTextRun,
} from "./index.js";

type MutableRunProperties = {
  -readonly [K in keyof SourceRunProperties]?: SourceRunProperties[K];
};
type MutableEditableTextRunProperties = {
  -readonly [K in keyof EditableTextRunProperties]?: EditableTextRunProperties[K];
};
type MutableParagraphProperties = {
  -readonly [K in keyof SourceParagraphProperties]?: SourceParagraphProperties[K];
};
type MutableEditableParagraphProperties = {
  -readonly [K in keyof EditableParagraphProperties]?: EditableParagraphProperties[K];
};

const EDITABLE_TEXT_RUN_PROPERTY_VALIDATORS: {
  readonly [K in EditableTextRunProperty]: (value: EditableTextRunProperties[K]) => void;
} = {
  bold: (value) => requireBooleanOrUndefined(value, "bold"),
  italic: (value) => requireBooleanOrUndefined(value, "italic"),
  underline: (value) => requireBooleanOrUndefined(value, "underline"),
  fontSize: (value) => {
    if (value !== undefined && (!Number.isFinite(value) || value <= 0)) {
      throw new Error("updateTextRunProperties: fontSize must be a finite positive pt value");
    }
  },
  color: (value) => {
    if (value === undefined) return;
    if (value.kind !== "srgb") {
      throw new Error("updateTextRunProperties: only srgb text run color is supported");
    }
    if (!/^[0-9A-Fa-f]{6}$/.test(value.hex)) {
      throw new Error("updateTextRunProperties: srgb text run color must be a 6-digit hex value");
    }
  },
  typeface: (value) => {
    if (value !== undefined && value.trim() === "") {
      throw new Error("updateTextRunProperties: typeface must be a non-empty string");
    }
  },
};
const EDITABLE_TEXT_RUN_PROPERTIES = Object.keys(EDITABLE_TEXT_RUN_PROPERTY_VALIDATORS).filter(
  (property): property is EditableTextRunProperty =>
    property in EDITABLE_TEXT_RUN_PROPERTY_VALIDATORS,
);
const EDITABLE_TEXT_RUN_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_TEXT_RUN_PROPERTIES);
const EDITABLE_PARAGRAPH_PROPERTIES = [
  "align",
  "level",
  "bullet",
] as const satisfies readonly EditableParagraphProperty[];
const EDITABLE_PARAGRAPH_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_PARAGRAPH_PROPERTIES);
const TEXT_ALIGN_VALUES: ReadonlySet<SourceTextAlign> = new Set([
  "left",
  "center",
  "right",
  "justify",
]);
const AUTO_NUM_SCHEMES: ReadonlySet<SourceAutoNumScheme> = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);

export function findTextRunBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceTextRun | undefined {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      const run = findTextRunInShape(shape, handle);
      if (run !== undefined) return run;
    }
  }
  return undefined;
}

export function findParagraphBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceParagraph | undefined {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      const paragraph = findParagraphInShape(shape, handle);
      if (paragraph !== undefined) return paragraph;
    }
  }
  return undefined;
}

export function replaceTextRunPlainText(
  source: PptxSourceModel,
  handle: SourceHandle,
  text: string,
): PptxSourceModel {
  const result = mapMatchingTextRun(source, handle, (run) =>
    run.text === text ? run : { ...run, text },
  );

  if (!result.matched) {
    throw new Error(
      "replaceTextRunPlainText: text run handle was not found in PptxSourceModel source",
    );
  }
  if (!result.changed) return source;

  return {
    ...source,
    slides: result.slides,
    edits: [...(source.edits ?? []), { kind: "replaceTextRunPlainText", handle, text }],
  };
}

export function setTextRunProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  properties: EditableTextRunProperties,
): PptxSourceModel {
  return updateTextRunProperties(source, handle, {
    set: properties,
    clear: [],
  });
}

export function clearTextRunProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  properties: readonly EditableTextRunProperty[],
): PptxSourceModel {
  return updateTextRunProperties(source, handle, {
    set: {},
    clear: properties,
  });
}

export function setParagraphProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  properties: EditableParagraphProperties,
): PptxSourceModel {
  return updateParagraphProperties(source, handle, {
    set: properties,
    clear: [],
  });
}

export function clearParagraphProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  properties: readonly EditableParagraphProperty[],
): PptxSourceModel {
  return updateParagraphProperties(source, handle, {
    set: {},
    clear: properties,
  });
}

export function replaceParagraphPlainText(
  source: PptxSourceModel,
  handle: SourceHandle,
  text: string,
): PptxSourceModel {
  const result = mapMatchingParagraph(source, handle, (paragraph) => {
    const replacementHandle = createReplacementRunHandle(paragraph);
    return {
      ...paragraph,
      runs: [
        {
          kind: "textRun",
          text,
          ...(paragraph.runs[0]?.properties !== undefined
            ? { properties: paragraph.runs[0].properties }
            : {}),
          ...(replacementHandle !== undefined ? { handle: replacementHandle } : {}),
        },
      ],
    } satisfies SourceParagraph;
  });

  if (!result.matched) {
    throw new Error(
      "replaceParagraphPlainText: paragraph handle was not found in PptxSourceModel source",
    );
  }

  return {
    ...source,
    slides: result.slides,
    edits: [...(source.edits ?? []), { kind: "replaceParagraphPlainText", handle, text }],
  };
}

interface UpdateTextRunPropertiesPatch {
  readonly set: EditableTextRunProperties;
  readonly clear: readonly EditableTextRunProperty[];
}

interface UpdateParagraphPropertiesPatch {
  readonly set: EditableParagraphProperties;
  readonly clear: readonly EditableParagraphProperty[];
}

function updateTextRunProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  patch: UpdateTextRunPropertiesPatch,
): PptxSourceModel {
  assertEditableTextRunProperties(patch.set);
  assertEditableTextRunPropertyNames(patch.clear);
  const set = definedEditableTextRunProperties(patch.set);
  if (Object.values(set).every((value) => value === undefined) && patch.clear.length === 0) {
    throw new Error("updateTextRunProperties: patch must set or clear at least one property");
  }

  const result = mapMatchingTextRun(source, handle, (run) => {
    const properties = patchTextRunProperties(run.properties, { set, clear: patch.clear });
    if (textRunPropertiesEqual(run.properties, properties)) return run;
    return {
      kind: run.kind,
      text: run.text,
      ...(run.handle !== undefined ? { handle: run.handle } : {}),
      ...(run.rawSidecars !== undefined ? { rawSidecars: run.rawSidecars } : {}),
      ...(properties !== undefined ? { properties } : {}),
    } satisfies SourceTextRun;
  });

  if (!result.matched) {
    throw new Error(
      "updateTextRunProperties: text run handle was not found in PptxSourceModel source",
    );
  }
  if (!result.changed) return source;

  return {
    ...source,
    slides: result.slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateTextRunProperties",
        handle,
        ...(Object.keys(set).length > 0 ? { set } : {}),
        ...(patch.clear.length > 0 ? { clear: patch.clear } : {}),
      },
    ],
  };
}

function updateParagraphProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  patch: UpdateParagraphPropertiesPatch,
): PptxSourceModel {
  assertEditableParagraphProperties(patch.set);
  assertEditableParagraphPropertyNames(patch.clear);
  const set = definedEditableParagraphProperties(patch.set);
  if (Object.values(set).every((value) => value === undefined) && patch.clear.length === 0) {
    throw new Error("updateParagraphProperties: patch must set or clear at least one property");
  }

  const result = mapMatchingParagraph(source, handle, (paragraph) => {
    const properties = patchParagraphProperties(paragraph.properties, { set, clear: patch.clear });
    if (paragraphPropertiesEqual(paragraph.properties, properties)) return paragraph;
    return {
      runs: paragraph.runs,
      ...(paragraph.handle !== undefined ? { handle: paragraph.handle } : {}),
      ...(paragraph.rawSidecars !== undefined ? { rawSidecars: paragraph.rawSidecars } : {}),
      ...(properties !== undefined ? { properties } : {}),
    } satisfies SourceParagraph;
  });

  if (!result.matched) {
    throw new Error(
      "updateParagraphProperties: paragraph handle was not found in PptxSourceModel source",
    );
  }
  if (!result.changed) return source;

  return {
    ...source,
    slides: result.slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateParagraphProperties",
        handle,
        ...(Object.keys(set).length > 0 ? { set } : {}),
        ...(patch.clear.length > 0 ? { clear: patch.clear } : {}),
      },
    ],
  };
}

interface TextEditMappingResult {
  readonly slides: readonly SourceSlide[];
  readonly matched: boolean;
  readonly changed: boolean;
}

interface ParagraphMappingResult {
  readonly paragraph: SourceParagraph;
  readonly matched: boolean;
}

function mapMatchingTextRun(
  source: PptxSourceModel,
  handle: SourceHandle,
  mapRun: (run: SourceTextRun) => SourceTextRun,
): TextEditMappingResult {
  return mapTextBodyParagraphs(source, (paragraph) => {
    let matched = false;
    let changed = false;
    const runs = paragraph.runs.map((run) => {
      if (!sourceHandlesEqual(run.handle, handle)) return run;
      matched = true;
      const mapped = mapRun(run);
      if (mapped === run) return run;
      changed = true;
      return mapped;
    });

    return {
      paragraph: changed ? ({ ...paragraph, runs } satisfies SourceParagraph) : paragraph,
      matched,
    };
  });
}

function mapMatchingParagraph(
  source: PptxSourceModel,
  handle: SourceHandle,
  mapParagraph: (paragraph: SourceParagraph) => SourceParagraph,
): TextEditMappingResult {
  return mapTextBodyParagraphs(source, (paragraph) => {
    if (!sourceHandlesEqual(paragraph.handle, handle)) {
      return { paragraph, matched: false };
    }
    return { paragraph: mapParagraph(paragraph), matched: true };
  });
}

function mapTextBodyParagraphs(
  source: PptxSourceModel,
  mapParagraph: (paragraph: SourceParagraph) => ParagraphMappingResult,
): TextEditMappingResult {
  let matched = false;
  let changed = false;

  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const shapes = slide.shapes.map((shape) => {
      if (shape.kind !== "shape" || shape.textBody === undefined) return shape;

      let shapeChanged = false;
      const paragraphs = shape.textBody.paragraphs.map((paragraph) => {
        const result = mapParagraph(paragraph);
        if (result.matched) matched = true;
        if (result.paragraph === paragraph) return paragraph;
        changed = true;
        shapeChanged = true;
        slideChanged = true;
        return result.paragraph;
      });

      if (!shapeChanged) return shape;
      return {
        ...shape,
        textBody: {
          ...shape.textBody,
          paragraphs,
        },
      } satisfies SourceShape;
    });

    return slideChanged ? { ...slide, shapes } : slide;
  });

  return {
    slides: changed ? slides : source.slides,
    matched,
    changed,
  };
}

function patchTextRunProperties(
  current: SourceRunProperties | undefined,
  patch: UpdateTextRunPropertiesPatch,
): SourceRunProperties | undefined {
  const next: MutableRunProperties = { ...(current ?? {}) };
  for (const property of patch.clear) {
    delete next[property];
  }
  Object.assign(next, patch.set);
  return Object.keys(next).length > 0 ? next : undefined;
}

function patchParagraphProperties(
  current: SourceParagraphProperties | undefined,
  patch: UpdateParagraphPropertiesPatch,
): SourceParagraphProperties | undefined {
  const next: MutableParagraphProperties = { ...(current ?? {}) };
  for (const property of patch.clear) {
    delete next[property];
  }
  Object.assign(next, patch.set);
  return Object.keys(next).length > 0 ? next : undefined;
}

function findTextRunInShape(shape: SourceShape, handle: SourceHandle): SourceTextRun | undefined {
  for (const paragraph of shape.textBody?.paragraphs ?? []) {
    for (const run of paragraph.runs) {
      if (sourceHandlesEqual(run.handle, handle)) return run;
    }
  }
  return undefined;
}

function findParagraphInShape(
  shape: SourceShape,
  handle: SourceHandle,
): SourceParagraph | undefined {
  return shape.textBody?.paragraphs.find((paragraph) =>
    sourceHandlesEqual(paragraph.handle, handle),
  );
}

function assertEditableTextRunProperties(properties: EditableTextRunProperties): void {
  for (const property of Object.keys(properties)) {
    assertEditableTextRunPropertyName(property);
  }
  for (const property of EDITABLE_TEXT_RUN_PROPERTIES) {
    validateEditableTextRunProperty(property, properties[property]);
  }
}

function assertEditableParagraphProperties(properties: EditableParagraphProperties): void {
  for (const property of Object.keys(properties)) {
    assertEditableParagraphPropertyName(property);
  }
  validateParagraphAlign(properties.align);
  validateParagraphLevel(properties.level);
  validateParagraphBullet(properties.bullet);
}

function requireBooleanOrUndefined(
  value: unknown,
  fieldName: "bold" | "italic" | "underline",
): void {
  if (value !== undefined && typeof value !== "boolean") {
    throw new Error(`updateTextRunProperties: ${fieldName} must be a boolean value`);
  }
}

function assertEditableTextRunPropertyNames(properties: readonly EditableTextRunProperty[]): void {
  for (const property of properties) {
    assertEditableTextRunPropertyName(property);
  }
}

function assertEditableParagraphPropertyNames(
  properties: readonly EditableParagraphProperty[],
): void {
  for (const property of properties) {
    assertEditableParagraphPropertyName(property);
  }
}

function definedEditableTextRunProperties(
  properties: EditableTextRunProperties,
): EditableTextRunProperties {
  const defined: MutableEditableTextRunProperties = {};
  for (const property of EDITABLE_TEXT_RUN_PROPERTIES) {
    copyDefinedEditableTextRunProperty(defined, properties, property);
  }
  return defined;
}

function definedEditableParagraphProperties(
  properties: EditableParagraphProperties,
): EditableParagraphProperties {
  const defined: MutableEditableParagraphProperties = {};
  for (const property of EDITABLE_PARAGRAPH_PROPERTIES) {
    copyDefinedEditableParagraphProperty(defined, properties, property);
  }
  return defined;
}

function assertEditableTextRunPropertyName(
  property: string,
): asserts property is EditableTextRunProperty {
  if (!isEditableTextRunProperty(property)) {
    throw new Error(`updateTextRunProperties: unsupported text run property '${property}'`);
  }
}

function assertEditableParagraphPropertyName(
  property: string,
): asserts property is EditableParagraphProperty {
  if (!EDITABLE_PARAGRAPH_PROPERTY_SET.has(property)) {
    throw new Error(`updateParagraphProperties: unsupported paragraph property '${property}'`);
  }
}

function isEditableTextRunProperty(property: string): property is EditableTextRunProperty {
  return EDITABLE_TEXT_RUN_PROPERTY_SET.has(property);
}

function validateParagraphAlign(value: EditableParagraphProperties["align"]): void {
  if (value !== undefined && !TEXT_ALIGN_VALUES.has(value)) {
    throw new Error("updateParagraphProperties: align must be left, center, right, or justify");
  }
}

function validateParagraphLevel(value: EditableParagraphProperties["level"]): void {
  if (value !== undefined && (!Number.isInteger(value) || value < 0 || value > 8)) {
    throw new Error("updateParagraphProperties: level must be an integer from 0 to 8");
  }
}

function validateParagraphBullet(value: EditableParagraphProperties["bullet"]): void {
  if (value === undefined) return;
  if (value.type === "none") return;
  if (value.type === "char") {
    if (typeof value.char !== "string" || value.char.length === 0) {
      throw new Error("updateParagraphProperties: bullet.char must be a non-empty string");
    }
    return;
  }
  if (value.type === "autoNum") {
    if (!AUTO_NUM_SCHEMES.has(value.scheme)) {
      throw new Error("updateParagraphProperties: unsupported bullet auto-numbering scheme");
    }
    if (!Number.isInteger(value.startAt) || value.startAt < 1) {
      throw new Error("updateParagraphProperties: bullet.startAt must be a positive integer");
    }
    return;
  }
  throw new Error("updateParagraphProperties: unsupported bullet type");
}

function validateEditableTextRunProperty<K extends EditableTextRunProperty>(
  property: K,
  value: EditableTextRunProperties[K],
): void {
  EDITABLE_TEXT_RUN_PROPERTY_VALIDATORS[property](value);
}

function copyDefinedEditableParagraphProperty<K extends EditableParagraphProperty>(
  target: MutableEditableParagraphProperties,
  source: EditableParagraphProperties,
  property: K,
): void {
  const value = source[property];
  if (value !== undefined) {
    target[property] = value;
  }
}

function copyDefinedEditableTextRunProperty<K extends EditableTextRunProperty>(
  target: MutableEditableTextRunProperties,
  source: EditableTextRunProperties,
  property: K,
): void {
  const value = source[property];
  if (value !== undefined) {
    target[property] = value;
  }
}

function textRunPropertiesEqual(
  left: SourceRunProperties | undefined,
  right: SourceRunProperties | undefined,
): boolean {
  return stableValueEqual(left ?? {}, right ?? {});
}

function paragraphPropertiesEqual(
  left: SourceParagraphProperties | undefined,
  right: SourceParagraphProperties | undefined,
): boolean {
  return stableValueEqual(left ?? {}, right ?? {});
}

function stableValueEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    if (left.length !== right.length) return false;
    return left.every((value, index) => stableValueEqual(value, right[index]));
  }
  if (isPlainRecord(left) || isPlainRecord(right)) {
    if (!isPlainRecord(left) || !isPlainRecord(right)) return false;
    const leftKeys = Object.keys(left).sort();
    const rightKeys = Object.keys(right).sort();
    if (!stableValueEqual(leftKeys, rightKeys)) return false;
    return leftKeys.every((key) => stableValueEqual(left[key], right[key]));
  }
  return false;
}

function isPlainRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function createReplacementRunHandle(paragraph: SourceParagraph): SourceHandle | undefined {
  if (paragraph.runs[0]?.handle !== undefined) return paragraph.runs[0].handle;
  if (paragraph.handle?.nodeId === undefined) return undefined;
  return {
    ...paragraph.handle,
    nodeId: asSourceNodeId(`${paragraph.handle.nodeId}:r:0`),
    orderingSlot: 0,
  };
}
