import type {
  PptxSourceModelEdit,
  PptxSourceModelTextRunEdit,
  PptxSourceModelTextRunPropertiesEdit,
} from "../source/index.js";

export function validateEdits(edits: readonly PptxSourceModelEdit[]): void {
  const addedShapeKeys = new Set<string>();
  const runKeys = new Set<string>();
  const paragraphKeys = new Set<string>();
  const shapeKeys = new Set<string>();
  const deletedShapeKeys = new Set<string>();
  const textRunEdits: PptxSourceModelTextRunEdit[] = [];
  const textRunPropertiesEdits: PptxSourceModelTextRunPropertiesEdit[] = [];

  for (const edit of edits) {
    switch (edit.kind) {
      case "addTextBox":
      case "addConnector": {
        const key = [edit.slidePartPath, edit.shapeId].join("\u0000");
        if (addedShapeKeys.has(key)) {
          throw new Error(`writePptx: conflicting shape additions for shape id '${edit.shapeId}'`);
        }
        addedShapeKeys.add(key);
        break;
      }
      case "replaceTextRunPlainText": {
        const key = editHandleNodeKey(edit);
        if (runKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting text run edits for handle '${edit.handle.nodeId}'`,
          );
        }
        runKeys.add(key);
        textRunEdits.push(edit);
        break;
      }
      case "updateTextRunProperties":
        textRunPropertiesEdits.push(edit);
        break;
      case "updateParagraphProperties":
        break;
      case "replaceParagraphPlainText": {
        const key = editHandleNodeKey(edit);
        if (paragraphKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting paragraph text edits for handle '${edit.handle.nodeId}'`,
          );
        }
        paragraphKeys.add(key);
        break;
      }
      case "updateShapeTransform": {
        const key = editHandleNodeKey(edit);
        if (shapeKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting shape transform edits for handle '${String(edit.handle.nodeId)}'`,
          );
        }
        shapeKeys.add(key);
        break;
      }
      case "deleteShape": {
        const key = editHandleNodeKey(edit);
        if (deletedShapeKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting shape delete edits for handle '${String(edit.handle.nodeId)}'`,
          );
        }
        deletedShapeKeys.add(key);
        break;
      }
      case "replaceImage":
      case "addEmptySlideFromLayout":
      case "duplicateSlide":
      case "moveSlide":
      case "deleteSlide":
        break;
    }
  }

  for (const runEdit of textRunEdits) {
    const paragraphKey = textRunParagraphEditKey(runEdit);
    if (paragraphKey !== undefined && paragraphKeys.has(paragraphKey)) {
      throw new Error(
        `writePptx: conflicting text run and paragraph edits for handle '${runEdit.handle.nodeId}'`,
      );
    }
  }
  for (const runPropertiesEdit of textRunPropertiesEdits) {
    const paragraphKey = textRunParagraphEditKey(runPropertiesEdit);
    if (paragraphKey !== undefined && paragraphKeys.has(paragraphKey)) {
      throw new Error(
        `writePptx: conflicting text run properties and paragraph edits for handle '${runPropertiesEdit.handle.nodeId}'`,
      );
    }
  }
}

function editHandleNodeKey(edit: {
  readonly handle: PptxSourceModelTextRunEdit["handle"];
}): string {
  return [edit.handle.partPath, edit.handle.nodeId ?? "", edit.handle.relationshipId ?? ""].join(
    "\u0000",
  );
}

function textRunParagraphEditKey(
  edit: PptxSourceModelTextRunEdit | PptxSourceModelTextRunPropertiesEdit,
): string | undefined {
  const nodeId = String(edit.handle.nodeId ?? "");
  const byShapeId = /^(text:shape:.+:p:\d+):r:\d+$/.exec(nodeId);
  const byShapeSlot = /^(text:shapeSlot:\d+:p:\d+):r:\d+$/.exec(nodeId);
  const paragraphNodeId = byShapeId?.[1] ?? byShapeSlot?.[1];
  if (paragraphNodeId === undefined) return undefined;
  return [edit.handle.partPath, paragraphNodeId, edit.handle.relationshipId ?? ""].join("\u0000");
}
