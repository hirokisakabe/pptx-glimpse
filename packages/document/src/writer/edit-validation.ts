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
  const shapeFillKeys = new Set<string>();
  const shapeOutlineKeys = new Set<string>();
  const pictureCropKeys = new Set<string>();
  const deletedShapeKeys = new Set<string>();
  const backgroundKeys = new Set<string>();
  const textRunEdits: PptxSourceModelTextRunEdit[] = [];
  const textRunPropertiesEdits: PptxSourceModelTextRunPropertiesEdit[] = [];

  for (const edit of edits) {
    switch (edit.kind) {
      case "addTextBox":
      case "addShape":
      case "addConnector":
      case "addPicture":
      case "addChart":
      case "addTable": {
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
      case "updateShapeFill": {
        const key = editHandleNodeKey(edit);
        if (shapeFillKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting shape fill edits for handle '${String(edit.handle.nodeId)}'`,
          );
        }
        shapeFillKeys.add(key);
        break;
      }
      case "updateShapeOutline": {
        const key = editHandleNodeKey(edit);
        if (shapeOutlineKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting shape outline edits for handle '${String(edit.handle.nodeId)}'`,
          );
        }
        shapeOutlineKeys.add(key);
        break;
      }
      case "updatePictureCrop": {
        const key = [edit.handle.partPath, edit.handle.nodeId ?? ""].join("\u0000");
        if (pictureCropKeys.has(key)) {
          throw new Error(
            `writePptx: conflicting picture crop edits for handle '${String(edit.handle.nodeId)}'`,
          );
        }
        pictureCropKeys.add(key);
        break;
      }
      case "updateTableCellProperties":
        break;
      case "reorderShapes": {
        if (new Set(edit.shapeIds).size !== edit.shapeIds.length) {
          throw new Error("writePptx: reordered shape ids contain a duplicate shape");
        }
        break;
      }
      case "moveShapes": {
        if (edit.shapeIds.length === 0) {
          throw new Error("writePptx: moved shape ids must not be empty");
        }
        if (new Set(edit.shapeIds).size !== edit.shapeIds.length) {
          throw new Error("writePptx: moved shape ids contain a duplicate shape");
        }
        if (edit.beforeShapeId !== undefined && edit.shapeIds.includes(edit.beforeShapeId)) {
          throw new Error("writePptx: move anchor must not be inside the moved block");
        }
        if (edit.crossParent !== true && edit.destinationParentGroupId !== undefined) {
          throw new Error("writePptx: same-parent move must not specify a destination parent");
        }
        if (edit.crossParent === true && edit.parentGroupId === edit.destinationParentGroupId) {
          throw new Error("writePptx: cross-parent move requires different containers");
        }
        break;
      }
      case "groupShapes": {
        if (new Set(edit.shapeIds).size !== edit.shapeIds.length) {
          throw new Error("writePptx: grouped shape ids contain a duplicate shape");
        }
        if (edit.shapeIds.length < 2) {
          throw new Error("writePptx: grouped shape ids must contain at least two shapes");
        }
        const key = [edit.targetPartPath, edit.groupId].join("\u0000");
        if (addedShapeKeys.has(key)) {
          throw new Error(`writePptx: conflicting shape additions for shape id '${edit.groupId}'`);
        }
        addedShapeKeys.add(key);
        break;
      }
      case "ungroupShape":
        break;
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
      case "updateChartData":
      case "updateScatterChartData":
      case "updateBubbleChartData":
      case "addSlideLayout":
      case "cloneSlideLayout":
      case "addEmptySlideFromLayout":
      case "duplicateSlide":
      case "moveSlide":
      case "deleteSlide":
      case "updateThemeScheme":
        break;
      case "setBackground":
      case "setSlideBackground": {
        validateBackgroundImageMetadata(edit);
        const targetPartPath =
          edit.kind === "setBackground" ? edit.targetPartPath : edit.slidePartPath;
        if (backgroundKeys.has(targetPartPath)) {
          throw new Error(
            `writePptx: conflicting background edits for drawing part '${targetPartPath}'`,
          );
        }
        backgroundKeys.add(targetPartPath);
        break;
      }
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

function validateBackgroundImageMetadata(edit: {
  readonly relationshipId?: unknown;
  readonly mediaPartPath?: unknown;
  readonly contentType?: unknown;
}): void {
  const definedCount = [edit.relationshipId, edit.mediaPartPath, edit.contentType].filter(
    (value) => value !== undefined,
  ).length;
  if (definedCount !== 0 && definedCount !== 3) {
    throw new Error(
      "writePptx: background image relationship, media part, and content type must be provided together",
    );
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
  const paragraphNodeId =
    /^(text:(?:shape:.+|shapeSlot:\d+|table:.+:row:\d+:cell:\d+):p:\d+):r:\d+$/.exec(nodeId)?.[1];
  if (paragraphNodeId === undefined) return undefined;
  return [edit.handle.partPath, paragraphNodeId, edit.handle.relationshipId ?? ""].join("\u0000");
}
