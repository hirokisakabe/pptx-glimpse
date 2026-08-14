import {
  asEmu,
  asOoxmlPercent,
  asPartPath,
  asSourceNodeId,
  createPptx,
  type SourceHandle,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession, type EditorCommand } from "../index.js";

const missingShapeHandle: SourceHandle = {
  partPath: asPartPath("ppt/slides/slide1.xml"),
  nodeId: asSourceNodeId("shape:missing"),
};
const missingSlideHandle: SourceHandle = {
  partPath: asPartPath("ppt/slides/missing.xml"),
};
const missingThemeHandle: SourceHandle = {
  partPath: asPartPath("ppt/theme/missing.xml"),
};

const REJECTED_COMMAND_BY_KIND = {
  replaceTextRunPlainText: {
    kind: "replaceTextRunPlainText",
    handle: missingShapeHandle,
    text: "replacement",
  },
  replaceParagraphPlainText: {
    kind: "replaceParagraphPlainText",
    handle: missingShapeHandle,
    text: "replacement",
  },
  setTextRunProperties: {
    kind: "setTextRunProperties",
    handle: missingShapeHandle,
    properties: { bold: true },
  },
  clearTextRunProperties: {
    kind: "clearTextRunProperties",
    handle: missingShapeHandle,
    properties: ["bold"],
  },
  setParagraphProperties: {
    kind: "setParagraphProperties",
    handle: missingShapeHandle,
    properties: { align: "center" },
  },
  clearParagraphProperties: {
    kind: "clearParagraphProperties",
    handle: missingShapeHandle,
    properties: ["align"],
  },
  moveShape: {
    kind: "moveShape",
    handle: missingShapeHandle,
    offsetX: asEmu(1),
    offsetY: asEmu(1),
  },
  resizeShape: {
    kind: "resizeShape",
    handle: missingShapeHandle,
    width: asEmu(1),
    height: asEmu(1),
  },
  setShapeTransform: {
    kind: "setShapeTransform",
    handle: missingShapeHandle,
    offsetX: asEmu(1),
    offsetY: asEmu(1),
    width: asEmu(1),
    height: asEmu(1),
  },
  setShapeFill: {
    kind: "setShapeFill",
    handle: missingShapeHandle,
    fill: { kind: "none" },
  },
  setShapeOutline: {
    kind: "setShapeOutline",
    handle: missingShapeHandle,
    outline: { fill: { kind: "none" } },
  },
  addTextBox: {
    kind: "addTextBox",
    slideHandle: missingSlideHandle,
    offsetX: asEmu(1),
    offsetY: asEmu(1),
    width: asEmu(1),
    height: asEmu(1),
  },
  addConnector: {
    kind: "addConnector",
    slideHandle: missingSlideHandle,
    preset: "straightConnector1",
    offsetX: asEmu(1),
    offsetY: asEmu(1),
    width: asEmu(1),
    height: asEmu(1),
  },
  deleteShape: { kind: "deleteShape", handle: missingShapeHandle },
  groupShapes: { kind: "groupShapes", shapeHandles: [] },
  moveShapes: {
    kind: "moveShapes",
    shapeHandles: [],
    destinationHandle: missingSlideHandle,
  },
  moveShapesAcrossSlides: {
    kind: "moveShapesAcrossSlides",
    shapeHandles: [],
    destinationSlideHandle: missingSlideHandle,
  },
  ungroupShape: { kind: "ungroupShape", groupHandle: missingShapeHandle },
  replaceImage: {
    kind: "replaceImage",
    handle: missingShapeHandle,
    bytes: new Uint8Array([0]),
  },
  setPictureCrop: {
    kind: "setPictureCrop",
    handle: missingShapeHandle,
    left: asOoxmlPercent(1),
  },
  clearPictureCrop: { kind: "clearPictureCrop", handle: missingShapeHandle },
  updateChartData: {
    kind: "updateChartData",
    handle: missingShapeHandle,
    series: [{ name: "Series", categories: ["A"], values: [1] }],
  },
  updateScatterChartData: {
    kind: "updateScatterChartData",
    handle: missingShapeHandle,
    series: [{ name: "Series", xValues: [1], yValues: [1] }],
  },
  updateBubbleChartData: {
    kind: "updateBubbleChartData",
    handle: missingShapeHandle,
    series: [{ name: "Series", xValues: [1], yValues: [1], bubbleSizes: [1] }],
  },
  addEmptySlideFromLayout: {
    kind: "addEmptySlideFromLayout",
    layoutPartPath: asPartPath("ppt/slideLayouts/missing.xml"),
  },
  duplicateSlide: { kind: "duplicateSlide", handle: missingSlideHandle },
  moveSlide: { kind: "moveSlide", handle: missingSlideHandle, toIndex: 0 },
  deleteSlide: { kind: "deleteSlide", handle: missingSlideHandle },
  updateThemeScheme: {
    kind: "updateThemeScheme",
    handle: missingThemeHandle,
    colorScheme: { accent1: "123456" },
  },
} satisfies {
  readonly [Kind in EditorCommand["kind"]]: Extract<EditorCommand, { readonly kind: Kind }>;
};

describe("editor command pipeline contract", () => {
  it("routes every EditorCommand kind through expected-rejection classification atomically", () => {
    for (const command of Object.values(REJECTED_COMMAND_BY_KIND)) {
      const source = createPptx();
      const session = createEditorSession(source);

      const result = session.apply(command);

      expect(result, command.kind).toMatchObject({ ok: false, code: "invalid-command" });
      expect(session.document, command.kind).toBe(source);
      expect(session.selection, command.kind).toBeUndefined();
      expect(session.undoDepth, command.kind).toBe(0);
      expect(session.redoDepth, command.kind).toBe(0);
    }
  });
});
