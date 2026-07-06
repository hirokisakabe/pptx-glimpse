import {
  asPartPath,
  asRelationshipId,
  asSourceNodeId,
  type EditableParagraphProperties,
  type EditableParagraphProperty,
  type SourceAutoNumScheme,
  type SourceHandle,
  type SourceParagraph,
  type SourceParagraphProperties,
  type SourceRunProperties,
  type SourceTextBody,
  type SourceTextRun,
} from "@pptx-glimpse/document";
import { Node as ProseMirrorNode, Schema } from "prosemirror-model";

export const pptxTextBodySchema = new Schema({
  nodes: {
    doc: { content: "paragraph*" },
    paragraph: {
      content: "text*",
      attrs: {
        handle: { default: null },
        properties: { default: null },
      },
      toDOM: () => ["p", 0],
    },
    text: { group: "inline" },
  },
  marks: {
    pptxRun: {
      attrs: {
        handle: { default: null },
        properties: { default: null },
      },
      toDOM: () => ["span", { "data-pptx-run": "" }, 0],
    },
  },
});

export interface PptxTextBodyProseMirrorDocJson {
  readonly type: "doc";
  readonly content?: readonly PptxTextBodyProseMirrorParagraphJson[];
}

export interface PptxTextBodyProseMirrorParagraphJson {
  readonly type: "paragraph";
  readonly attrs?: {
    readonly handle?: SourceHandle | null;
    /** Source metadata plus editable paragraph properties. Only align / level / bullet are applied. */
    readonly properties?: unknown;
  };
  readonly content?: readonly PptxTextBodyProseMirrorTextJson[];
}

export interface PptxTextBodyProseMirrorTextJson {
  readonly type: "text";
  readonly text: string;
  readonly marks?: readonly PptxTextBodyProseMirrorRunMarkJson[];
}

export interface PptxTextBodyProseMirrorRunMarkJson {
  readonly type: "pptxRun";
  /** Read-only source metadata. Formatting edits are not applied from ProseMirror JSON. */
  readonly attrs?: {
    readonly handle?: SourceHandle | null;
    readonly properties?: unknown;
  };
}

export type PptxTextBodyProseMirrorCommand =
  | {
      readonly kind: "replaceTextRunPlainText";
      readonly handle: SourceHandle;
      readonly text: string;
    }
  | {
      readonly kind: "replaceParagraphPlainText";
      readonly handle: SourceHandle;
      readonly text: string;
    }
  | {
      readonly kind: "setParagraphProperties";
      readonly handle: SourceHandle;
      readonly properties: EditableParagraphProperties;
    }
  | {
      readonly kind: "clearParagraphProperties";
      readonly handle: SourceHandle;
      readonly properties: readonly EditableParagraphProperty[];
    };

interface RunGroup {
  readonly handle?: SourceHandle;
  readonly properties?: SourceRunProperties;
  readonly text: string;
}

const EDITABLE_PARAGRAPH_PROPERTIES = [
  "align",
  "level",
  "bullet",
] as const satisfies readonly EditableParagraphProperty[];
const AUTO_NUM_SCHEMES = [
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
] as const satisfies readonly SourceAutoNumScheme[];
const AUTO_NUM_SCHEME_SET: ReadonlySet<string> = new Set(AUTO_NUM_SCHEMES);

export function textBodyToProseMirrorDocJson(
  textBody: SourceTextBody,
): PptxTextBodyProseMirrorDocJson {
  assertSupportedTextBody(textBody);
  return {
    type: "doc",
    content: textBody.paragraphs.map((paragraph) => {
      const content: PptxTextBodyProseMirrorTextJson[] = paragraph.runs.map((run) => ({
        type: "text",
        text: run.text,
        marks: [
          {
            type: "pptxRun",
            attrs: {
              handle: run.handle ?? null,
              properties: run.properties ?? null,
            },
          },
        ],
      }));

      return {
        type: "paragraph",
        attrs: {
          handle: paragraph.handle ?? null,
          properties: paragraph.properties ?? null,
        },
        ...(content.length > 0 ? { content } : {}),
      };
    }),
  };
}

function assertSupportedTextBody(textBody: SourceTextBody): void {
  textBody.paragraphs.forEach((paragraph, paragraphIndex) => {
    paragraph.runs.forEach((run, runIndex) => {
      if (run.text.length === 0) {
        throw new Error(
          `textBodyToProseMirrorDocJson: empty text runs are unsupported at paragraph ${paragraphIndex}, run ${runIndex}`,
        );
      }
      if (run.text.includes("\n") || (run.rawSidecars?.length ?? 0) > 0) {
        throw new Error(
          `textBodyToProseMirrorDocJson: unsupported run-like content at paragraph ${paragraphIndex}, run ${runIndex}`,
        );
      }
    });
  });
}

export function proseMirrorDocJsonToTextBody(
  originalTextBody: SourceTextBody,
  docJson: unknown,
): SourceTextBody {
  const parsed = parsePptxTextBodyProseMirrorDocJson(docJson);
  validateProseMirrorDocJson(parsed);
  const paragraphs = (parsed.content ?? []).map((paragraph, index) =>
    paragraphJsonToSourceParagraph(originalTextBody.paragraphs[index], paragraph),
  );

  return {
    ...originalTextBody,
    paragraphs,
  };
}

export function proseMirrorDocJsonToEditorCommands(
  originalTextBody: SourceTextBody,
  docJson: unknown,
): readonly PptxTextBodyProseMirrorCommand[] {
  const editedTextBody = proseMirrorDocJsonToTextBody(originalTextBody, docJson);
  if (editedTextBody.paragraphs.length !== originalTextBody.paragraphs.length) {
    throw new Error("proseMirrorDocJsonToEditorCommands: paragraph count changes are unsupported");
  }

  return editedTextBody.paragraphs.flatMap<PptxTextBodyProseMirrorCommand>(
    (editedParagraph, paragraphIndex) => {
      const originalParagraph = originalTextBody.paragraphs[paragraphIndex];
      if (originalParagraph === undefined) return [];
      const paragraphPropertyCommands = paragraphPropertiesToEditorCommands(
        originalParagraph,
        editedParagraph,
      );
      if (paragraphRunHandlesMatch(originalParagraph, editedParagraph)) {
        const textRunCommands = editedParagraph.runs.flatMap<PptxTextBodyProseMirrorCommand>(
          (editedRun, runIndex) => {
            const originalRun = originalParagraph.runs[runIndex];
            if (originalRun === undefined || editedRun.text === originalRun.text) return [];
            if (editedRun.handle === undefined) {
              throw new Error(
                "proseMirrorDocJsonToEditorCommands: changed text run has no source handle",
              );
            }
            return [
              {
                kind: "replaceTextRunPlainText",
                handle: editedRun.handle,
                text: editedRun.text,
              },
            ];
          },
        );
        return [...paragraphPropertyCommands, ...textRunCommands];
      }

      if (editedParagraph.handle === undefined) {
        throw new Error(
          "proseMirrorDocJsonToEditorCommands: paragraph structure changed but paragraph has no source handle",
        );
      }
      return [
        ...paragraphPropertyCommands,
        {
          kind: "replaceParagraphPlainText",
          handle: editedParagraph.handle,
          text: paragraphPlainText(editedParagraph),
        },
      ];
    },
  );
}

function validateProseMirrorDocJson(docJson: PptxTextBodyProseMirrorDocJson): void {
  const doc = ProseMirrorNode.fromJSON(pptxTextBodySchema, docJson);
  doc.check();
}

function paragraphJsonToSourceParagraph(
  originalParagraph: SourceParagraph | undefined,
  paragraphJson: PptxTextBodyProseMirrorParagraphJson,
): SourceParagraph {
  const groups = collectRunGroups(paragraphJson, originalParagraph);
  const properties = paragraphPropertiesFromJson(
    originalParagraph?.properties,
    paragraphJson.attrs?.properties,
  );

  return {
    runs: groups.map(sourceRunFromGroup),
    ...(originalParagraph?.rawSidecars !== undefined
      ? { rawSidecars: originalParagraph.rawSidecars }
      : {}),
    ...(properties !== undefined ? { properties } : {}),
    ...(originalParagraph?.handle !== undefined
      ? { handle: originalParagraph.handle }
      : sourceHandleFromUnknown(paragraphJson.attrs?.handle) !== undefined
        ? { handle: sourceHandleFromUnknown(paragraphJson.attrs?.handle) }
        : {}),
  };
}

function collectRunGroups(
  paragraphJson: PptxTextBodyProseMirrorParagraphJson,
  originalParagraph: SourceParagraph | undefined,
): readonly RunGroup[] {
  const groups: RunGroup[] = [];
  let previousHandle = originalParagraph?.runs[0]?.handle;

  for (const textNode of paragraphJson.content ?? []) {
    const mark = textNode.marks?.find((candidate) => candidate.type === "pptxRun");
    const markedHandle = sourceHandleFromUnknown(mark?.attrs?.handle);
    const handle = markedHandle ?? previousHandle;
    const originalRun =
      handle === undefined ? undefined : findOriginalRun(originalParagraph, handle);
    const properties = originalRun?.properties;
    const lastGroup = groups.at(-1);

    if (lastGroup !== undefined && sourceHandlesEqual(lastGroup.handle, handle)) {
      groups[groups.length - 1] = {
        ...lastGroup,
        text: lastGroup.text + textNode.text,
      };
    } else {
      groups.push({
        ...(handle !== undefined ? { handle } : {}),
        ...(properties !== undefined ? { properties } : {}),
        text: textNode.text,
      });
    }
    previousHandle = handle;
  }

  return groups;
}

function sourceRunFromGroup(group: RunGroup): SourceTextRun {
  return {
    kind: "textRun",
    text: group.text,
    ...(group.properties !== undefined ? { properties: group.properties } : {}),
    ...(group.handle !== undefined ? { handle: group.handle } : {}),
  };
}

function findOriginalRun(
  paragraph: SourceParagraph | undefined,
  handle: SourceHandle,
): SourceTextRun | undefined {
  return paragraph?.runs.find((run) => sourceHandlesEqual(run.handle, handle));
}

function paragraphRunHandlesMatch(
  originalParagraph: SourceParagraph,
  editedParagraph: SourceParagraph,
): boolean {
  if (originalParagraph.runs.length !== editedParagraph.runs.length) return false;
  return originalParagraph.runs.every((run, index) =>
    sourceHandlesEqual(run.handle, editedParagraph.runs[index]?.handle),
  );
}

function paragraphPlainText(paragraph: SourceParagraph): string {
  return paragraph.runs.map((run) => run.text).join("");
}

function paragraphPropertiesToEditorCommands(
  originalParagraph: SourceParagraph,
  editedParagraph: SourceParagraph,
): readonly PptxTextBodyProseMirrorCommand[] {
  if (editedParagraph.handle === undefined) {
    if (
      paragraphEditablePropertiesEqual(originalParagraph.properties, editedParagraph.properties)
    ) {
      return [];
    }
    throw new Error("proseMirrorDocJsonToEditorCommands: changed paragraph has no source handle");
  }

  const set: MutableEditableParagraphProperties = {};
  const clear: EditableParagraphProperty[] = [];
  if (!stableValueEqual(originalParagraph.properties?.align, editedParagraph.properties?.align)) {
    if (editedParagraph.properties?.align === undefined) clear.push("align");
    else set.align = editedParagraph.properties.align;
  }
  if (!stableValueEqual(originalParagraph.properties?.level, editedParagraph.properties?.level)) {
    if (editedParagraph.properties?.level === undefined) clear.push("level");
    else set.level = editedParagraph.properties.level;
  }
  if (!stableValueEqual(originalParagraph.properties?.bullet, editedParagraph.properties?.bullet)) {
    if (editedParagraph.properties?.bullet === undefined) clear.push("bullet");
    else set.bullet = editedParagraph.properties.bullet;
  }

  const commands: PptxTextBodyProseMirrorCommand[] = [];
  if (clear.length > 0) {
    commands.push({
      kind: "clearParagraphProperties",
      handle: editedParagraph.handle,
      properties: clear,
    });
  }
  if (Object.keys(set).length > 0) {
    commands.push({
      kind: "setParagraphProperties",
      handle: editedParagraph.handle,
      properties: set,
    });
  }
  return commands;
}

function paragraphPropertiesFromJson(
  originalProperties: SourceParagraphProperties | undefined,
  value: unknown,
): SourceParagraphProperties | undefined {
  if (!isRecord(value)) return originalProperties;
  const next: MutableSourceParagraphProperties = { ...(originalProperties ?? {}) };
  for (const property of EDITABLE_PARAGRAPH_PROPERTIES) {
    if (!Object.prototype.hasOwnProperty.call(value, property)) {
      delete next[property];
      continue;
    }
    const propertyValue = value[property];
    if (propertyValue === null || propertyValue === undefined) {
      delete next[property];
    } else if (property === "align") {
      next.align = paragraphAlignFromUnknown(propertyValue);
    } else if (property === "level") {
      next.level = paragraphLevelFromUnknown(propertyValue);
    } else {
      next.bullet = paragraphBulletFromUnknown(propertyValue);
    }
  }
  return Object.keys(next).length > 0 ? next : undefined;
}

function paragraphAlignFromUnknown(value: unknown): EditableParagraphProperties["align"] {
  if (value === "left" || value === "center" || value === "right" || value === "justify") {
    return value;
  }
  throw new Error("ProseMirror paragraph properties align must be left, center, right, or justify");
}

function paragraphLevelFromUnknown(value: unknown): EditableParagraphProperties["level"] {
  if (typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= 8) {
    return value;
  }
  throw new Error("ProseMirror paragraph properties level must be an integer from 0 to 8");
}

function paragraphBulletFromUnknown(value: unknown): EditableParagraphProperties["bullet"] {
  if (!isRecord(value) || typeof value.type !== "string") {
    throw new Error("ProseMirror paragraph properties bullet must be an object");
  }
  if (value.type === "none") return { type: "none" };
  if (value.type === "char") {
    if (typeof value.char !== "string" || value.char.length === 0) {
      throw new Error("ProseMirror paragraph properties bullet.char must be a non-empty string");
    }
    return { type: "char", char: value.char };
  }
  if (value.type === "autoNum") {
    if (!isAutoNumScheme(value.scheme)) {
      throw new Error("ProseMirror paragraph properties bullet.scheme is unsupported");
    }
    if (
      typeof value.startAt !== "number" ||
      !Number.isInteger(value.startAt) ||
      value.startAt < 1
    ) {
      throw new Error("ProseMirror paragraph properties bullet.startAt must be positive");
    }
    return { type: "autoNum", scheme: value.scheme, startAt: value.startAt };
  }
  throw new Error("ProseMirror paragraph properties bullet.type is unsupported");
}

function isAutoNumScheme(value: unknown): value is SourceAutoNumScheme {
  return typeof value === "string" && AUTO_NUM_SCHEME_SET.has(value);
}

function paragraphEditablePropertiesEqual(
  left: SourceParagraphProperties | undefined,
  right: SourceParagraphProperties | undefined,
): boolean {
  return EDITABLE_PARAGRAPH_PROPERTIES.every((property) =>
    stableValueEqual(left?.[property], right?.[property]),
  );
}

function parsePptxTextBodyProseMirrorDocJson(value: unknown): PptxTextBodyProseMirrorDocJson {
  if (!isRecord(value) || value.type !== "doc") {
    throw new Error("ProseMirror text body doc JSON must be a doc node");
  }
  if (value.content !== undefined && readArray(value.content, isParagraphJson) === undefined) {
    throw new Error("ProseMirror text body doc JSON content must contain paragraph nodes");
  }
  const content = readArray(value.content, isParagraphJson)?.map(normalizeParagraphJson);
  return {
    type: "doc",
    ...(content !== undefined ? { content } : {}),
  };
}

function isParagraphJson(value: unknown): value is PptxTextBodyProseMirrorParagraphJson {
  if (!isRecord(value) || value.type !== "paragraph") return false;
  if (value.content !== undefined && readArray(value.content, isTextJson) === undefined) {
    return false;
  }
  return value.attrs === undefined || isRecord(value.attrs);
}

function normalizeParagraphJson(
  value: PptxTextBodyProseMirrorParagraphJson,
): PptxTextBodyProseMirrorParagraphJson {
  return {
    type: "paragraph",
    ...(value.attrs !== undefined ? { attrs: value.attrs } : {}),
    ...(value.content !== undefined ? { content: value.content } : {}),
  };
}

function isTextJson(value: unknown): value is PptxTextBodyProseMirrorTextJson {
  if (!isRecord(value) || value.type !== "text" || typeof value.text !== "string") return false;
  if (value.marks === undefined) return true;
  return readArray(value.marks, isRunMarkJson) !== undefined;
}

function isRunMarkJson(value: unknown): value is PptxTextBodyProseMirrorRunMarkJson {
  if (!isRecord(value) || value.type !== "pptxRun") return false;
  return value.attrs === undefined || isRecord(value.attrs);
}

function readArray<T>(
  value: unknown,
  itemGuard: (item: unknown) => item is T,
): readonly T[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) return undefined;
  return value.every(itemGuard) ? value : undefined;
}

function sourceHandleFromUnknown(value: unknown): SourceHandle | undefined {
  if (value === null || value === undefined) return undefined;
  if (!isRecord(value) || typeof value.partPath !== "string") return undefined;
  const handle = {
    partPath: asPartPath(value.partPath),
    ...(typeof value.nodeId === "string" ? { nodeId: asSourceNodeId(value.nodeId) } : {}),
    ...(typeof value.relationshipId === "string"
      ? { relationshipId: asRelationshipId(value.relationshipId) }
      : {}),
    ...(typeof value.orderingSlot === "number" ? { orderingSlot: value.orderingSlot } : {}),
  } satisfies SourceHandle;
  return handle;
}

function sourceHandlesEqual(
  left: SourceHandle | undefined,
  right: SourceHandle | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return sourceHandleKey(left) === sourceHandleKey(right);
}

function sourceHandleKey(handle: SourceHandle): string {
  return [
    handle.partPath,
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot ?? "",
  ].join("\u0000");
}

function stableValueEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    if (left.length !== right.length) return false;
    return left.every((value, index) => stableValueEqual(value, right[index]));
  }
  if (isRecord(left) || isRecord(right)) {
    if (!isRecord(left) || !isRecord(right)) return false;
    const leftKeys = Object.keys(left).sort();
    const rightKeys = Object.keys(right).sort();
    if (!stableValueEqual(leftKeys, rightKeys)) return false;
    return leftKeys.every((key) => stableValueEqual(left[key], right[key]));
  }
  return false;
}

function isRecord(value: unknown): value is { readonly [key: string]: unknown } {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type MutableSourceParagraphProperties = {
  -readonly [K in keyof SourceParagraphProperties]?: SourceParagraphProperties[K];
};

type MutableEditableParagraphProperties = {
  -readonly [K in keyof EditableParagraphProperties]?: EditableParagraphProperties[K];
};
