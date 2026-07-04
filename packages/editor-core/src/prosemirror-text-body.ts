import {
  asPartPath,
  asRelationshipId,
  asSourceNodeId,
  type SourceHandle,
  type SourceParagraph,
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
  /** Read-only source metadata. Formatting edits are not applied from ProseMirror JSON. */
  readonly attrs?: {
    readonly handle?: SourceHandle | null;
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
    };

interface RunGroup {
  readonly handle?: SourceHandle;
  readonly properties?: SourceRunProperties;
  readonly text: string;
}

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
      if (paragraphRunHandlesMatch(originalParagraph, editedParagraph)) {
        return editedParagraph.runs.flatMap<PptxTextBodyProseMirrorCommand>(
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
      }

      if (editedParagraph.handle === undefined) {
        throw new Error(
          "proseMirrorDocJsonToEditorCommands: paragraph structure changed but paragraph has no source handle",
        );
      }
      return [
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

  return {
    ...(originalParagraph ?? {}),
    runs: groups.map(sourceRunFromGroup),
    ...(originalParagraph?.properties !== undefined
      ? { properties: originalParagraph.properties }
      : {}),
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

function parsePptxTextBodyProseMirrorDocJson(value: unknown): PptxTextBodyProseMirrorDocJson {
  if (!isRecord(value) || value.type !== "doc") {
    throw new Error("ProseMirror text body doc JSON must be a doc node");
  }
  if (value.content !== undefined && readArray(value.content, isParagraphJson) === undefined) {
    throw new Error("ProseMirror text body doc JSON content must contain paragraph nodes");
  }
  const content = readArray(value.content, isParagraphJson);
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

function isRecord(value: unknown): value is { readonly [key: string]: unknown } {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
