import { asOoxmlPercent, asPartPath, readPptx, writePptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";
import {
  BLUE_PNG,
  buildBubbleChartEditSource,
  buildCategoryComboChartEditSource,
  buildChartEditSource,
  buildImageReplacementFixture,
  buildScatterChartEditSource,
  buildTwoSlideSharedImageFixture,
  chartXml,
  expectApplied,
  expectHistory,
  firstChart,
  firstImage,
  firstShape,
  JPEG_BYTES,
  mediaBytes,
  RED_PNG,
  requireHandle,
} from "./index.test-helpers.js";

describe("EditorSession image replacement commands", () => {
  it("replaces a pic image media part and persists it through write/read", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const session = createEditorSession(source);
    const image = firstImage(source);
    const result = session.apply({
      kind: "replaceImage",
      handle: requireHandle(image.handle),
      bytes: BLUE_PNG,
    });
    const edited = expectApplied(result);
    const reread = readPptx(writePptx(edited));

    expect(mediaBytes(source, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(edited, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(edited, "ppt/media/image2.png")).toEqual(BLUE_PNG);
    expect(mediaBytes(reread, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(reread, "ppt/media/image2.png")).toEqual(BLUE_PNG);
    expect(result.ok && result.warnings).toBeUndefined();
    expect(firstImage(reread).blipRelationshipId).toBe("rId1");
    expect(
      reread.slides[0]?.shapes.find(
        (shape) => shape.kind === "image" && shape.name === "Shared Picture B",
      ),
    ).toMatchObject({ blipRelationshipId: "rIdImage" });
  });

  it("keeps the other shared picture unchanged when a later batch command deletes the target slide", async () => {
    const source = readPptx(await buildTwoSlideSharedImageFixture());
    const session = createEditorSession(source);
    const result = session.applyAll([
      {
        kind: "replaceImage",
        handle: requireHandle(firstImage(source).handle),
        bytes: BLUE_PNG,
      },
      {
        kind: "deleteSlide",
        handle: requireHandle(source.slides[0]?.handle),
      },
    ]);
    const edited = expectApplied(result);

    expect(edited.slides).toHaveLength(1);
    expect(edited.slides[0]?.partPath).toBe(asPartPath("ppt/slides/slide2.xml"));
    expect(mediaBytes(edited, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(result.ok && result.warnings).toBeUndefined();
  });

  it("reuses the copy-on-write media for repeated replacements in one batch", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const session = createEditorSession(source);
    const imageHandle = requireHandle(firstImage(source).handle);
    const result = session.applyAll([
      { kind: "replaceImage", handle: imageHandle, bytes: BLUE_PNG },
      { kind: "replaceImage", handle: imageHandle, bytes: RED_PNG },
    ]);

    expectApplied(result);
    expect(result.ok && result.warnings).toBeUndefined();
    expect(mediaBytes(session.document, "ppt/media/image2.png")).toEqual(RED_PNG);
  });

  it("undoes and redoes image replacement by restoring media bytes", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const session = createEditorSession(source);

    expectApplied(
      session.apply({
        kind: "replaceImage",
        handle: requireHandle(firstImage(source).handle),
        bytes: BLUE_PNG,
      }),
    );

    const undone = expectHistory(session.undo());
    const redone = expectHistory(session.redo());

    expect(mediaBytes(undone, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(readPptx(writePptx(undone)), "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(redone, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(redone, "ppt/media/image2.png")).toEqual(BLUE_PNG);
    expect(mediaBytes(readPptx(writePptx(redone)), "ppt/media/image2.png")).toEqual(BLUE_PNG);
  });

  it("rejects invalid image replacement commands without changing document state", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const session = createEditorSession(source);
    const before = session.document;
    const imageHandle = requireHandle(firstImage(source).handle);
    const shapeHandle = requireHandle(firstShape(source).handle);

    const nonPic = session.apply({ kind: "replaceImage", handle: shapeHandle, bytes: BLUE_PNG });
    const unknownFormat = session.apply({
      kind: "replaceImage",
      handle: imageHandle,
      bytes: new Uint8Array([1, 2, 3]),
    });
    const differentFormat = session.apply({
      kind: "replaceImage",
      handle: imageHandle,
      bytes: JPEG_BYTES,
    });

    expect(nonPic).toMatchObject({ ok: false, code: "invalid-command" });
    expect(unknownFormat).toMatchObject({ ok: false, code: "invalid-command" });
    expect(differentFormat).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(mediaBytes(session.document, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(session.undoDepth).toBe(0);
  });
});

describe("EditorSession picture crop commands", () => {
  it("sets, clears, and restores crop through commands and history", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const image = firstImage(source);
    const imageHandle = requireHandle(image.handle);
    const session = createEditorSession(source);

    const setResult = session.setPictureCrop(image, {
      left: asOoxmlPercent(15000),
      bottom: asOoxmlPercent(5000),
    });
    const cropped = expectApplied(setResult);
    expect(firstImage(cropped).crop).toEqual({ left: 15000, bottom: 5000 });
    expect(firstImage(readPptx(writePptx(cropped))).crop).toEqual({
      left: 15000,
      bottom: 5000,
    });

    const clearResult = session.apply({ kind: "clearPictureCrop", handle: imageHandle });
    expect(firstImage(expectApplied(clearResult)).crop).toBeUndefined();
    expect(session.undoDepth).toBe(2);
    expect(firstImage(expectHistory(session.undo())).crop).toEqual({
      left: 15000,
      bottom: 5000,
    });
    expect(firstImage(expectHistory(session.undo())).crop).toBeUndefined();
    expect(firstImage(expectHistory(session.redo())).crop).toEqual({
      left: 15000,
      bottom: 5000,
    });
  });

  it("does not let extra crop input properties override the convenience command target", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const image = firstImage(source);
    const imageHandle = requireHandle(image.handle);
    const session = createEditorSession(source);
    const cropWithOverrides = {
      left: asOoxmlPercent(10000),
      kind: "clearPictureCrop" as const,
      handle: requireHandle(firstShape(source).handle),
    };

    const edited = expectApplied(session.setPictureCrop(image, cropWithOverrides));

    expect(firstImage(edited).crop).toEqual({ left: 10000 });
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "updatePictureCrop",
      handle: imageHandle,
    });
  });

  it("normalizes repeated crop commands and rejects invalid input atomically", async () => {
    const source = readPptx(await buildImageReplacementFixture());
    const imageHandle = requireHandle(firstImage(source).handle);
    const session = createEditorSession(source);
    const result = session.applyAll([
      { kind: "setPictureCrop", handle: imageHandle, left: asOoxmlPercent(10000) },
      { kind: "setPictureCrop", handle: imageHandle, right: asOoxmlPercent(20000) },
    ]);
    const edited = expectApplied(result);
    expect(firstImage(edited).crop).toEqual({ right: 20000 });
    expect(edited.edits?.filter((edit) => edit.kind === "updatePictureCrop")).toEqual([
      {
        kind: "updatePictureCrop",
        handle: imageHandle,
        crop: { right: 20000 },
      },
    ]);

    const invalidSession = createEditorSession(source);
    const before = invalidSession.document;
    const invalid = invalidSession.apply({
      kind: "setPictureCrop",
      handle: imageHandle,
      left: asOoxmlPercent(60000),
      right: asOoxmlPercent(40000),
    });
    expect(invalid).toMatchObject({ ok: false, code: "invalid-command" });
    expect(invalidSession.document).toBe(before);
    expect(invalidSession.undoDepth).toBe(0);
  });
});

describe("EditorSession chart data commands", () => {
  it("updates a category combo and restores it through undo and redo", async () => {
    const source = await buildCategoryComboChartEditSource();
    const chart = firstChart(source);
    const session = createEditorSession(source);
    const applied = session.apply({
      kind: "updateChartData",
      handle: requireHandle(chart.handle),
      series: [
        {
          source: { chartType: "bar", index: 0 },
          name: "Edited columns",
          categories: ["X", "Y", "Z"],
          values: [3, 5, 8],
        },
        {
          source: { chartType: "line", index: 7 },
          name: "Edited trend",
          categories: ["X", "Y", "Z"],
          values: [2, 4, 7],
        },
      ],
    });
    expectApplied(applied);
    expect(chartXml(session.document)).toContain("Edited columns");
    expect(chartXml(session.document)).toContain("Edited trend");
    expect(chartXml(expectHistory(session.undo()))).toContain("Original 1");
    expect(chartXml(expectHistory(session.redo()))).toContain("Edited trend");

    const before = session.document;
    const undoDepth = session.undoDepth;
    const rejected = session.apply({
      kind: "updateChartData",
      handle: requireHandle(chart.handle),
      series: [
        {
          source: { chartType: "bar", index: 0 },
          name: "Only",
          categories: ["A"],
          values: [1],
        },
      ],
    });
    expect(rejected).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(undoDepth);
  });

  it("rejects unsupported combo grouping and axis topology without changing editor state", async () => {
    const sources = [
      await buildCategoryComboChartEditSource({ lineGrouping: "stacked" }),
      await buildCategoryComboChartEditSource({ lineAxisIds: ["100003"] }),
    ];
    for (const source of sources) {
      const chart = firstChart(source);
      const handle = requireHandle(chart.handle);
      const session = createEditorSession(source);
      session.selectShape(handle);
      const before = session.document;
      const result = session.apply({
        kind: "updateChartData",
        handle,
        series: [
          {
            source: { chartType: "bar", index: 0 },
            name: "Columns",
            categories: ["A"],
            values: [1],
          },
          {
            source: { chartType: "line", index: 7 },
            name: "Trend",
            categories: ["A"],
            values: [2],
          },
        ],
      });

      expect(result).toMatchObject({ ok: false, code: "invalid-command" });
      expect(session.document).toBe(before);
      expect(session.selection).toEqual({ shapeHandle: handle });
      expect(session.undoDepth).toBe(0);
      expect(session.redoDepth).toBe(0);
    }
  });

  it("applies the convenience API and command with undo/redo history", () => {
    const source = buildChartEditSource();
    const chart = firstChart(source);
    const session = createEditorSession(source);
    const firstResult = session.updateChartData(chart, {
      series: [
        { name: "Edited 1", categories: ["A", "B", "C"], values: [3, 5, 8] },
        { name: "Edited 2", categories: ["A", "B", "C"], values: [2, 4, 6] },
        { name: "Edited 3", categories: ["A", "B", "C"], values: [1, 2, 3] },
      ],
    });
    expectApplied(firstResult);
    expect(chartXml(session.document)).toContain("Edited 1");
    expect(chartXml(session.document)).toContain("Sheet1!$B$2:$B$4");
    expect(chartXml(session.document)).toContain("Sheet1!$D$2:$D$4");
    expect(session.undoDepth).toBe(1);

    expect(chartXml(expectHistory(session.undo()))).not.toContain("Sheet1!$D$1");
    expect(chartXml(expectHistory(session.redo()))).toContain("Edited 3");

    const commandResult = session.apply({
      kind: "updateChartData",
      handle: requireHandle(chart.handle),
      series: [{ name: "Command 1", categories: ["X", "Y"], values: [10, 20] }],
    });
    expectApplied(commandResult);
    expect(chartXml(session.document)).not.toContain("Sheet1!$C$1");
    expect(session.undoDepth).toBe(2);
    expect(chartXml(expectHistory(session.undo()))).toContain("Edited 3");
    expect(chartXml(expectHistory(session.redo()))).not.toContain("Sheet1!$C$1");
  });

  it("rejects invalid chart updates without changing document, selection, or history", () => {
    const source = buildChartEditSource();
    const chart = firstChart(source);
    const session = createEditorSession(source);
    const handle = requireHandle(chart.handle);
    expect(session.selectShape(handle)).toMatchObject({ ok: true });
    const before = session.document;

    const result = session.apply({
      kind: "updateChartData",
      handle,
      series: [
        { name: "One", categories: ["A"], values: [1] },
        { name: "Two", categories: ["B"], values: [2] },
      ],
    });

    expect(result).toMatchObject({
      ok: false,
      code: "invalid-command",
    });
    expect(result.ok ? undefined : result.message).toContain(
      "every series must use identical category labels",
    );
    expect(session.document).toBe(before);
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });
});

describe("EditorSession scatter chart data commands", () => {
  it("applies the convenience API and command with undo/redo history", async () => {
    const source = await buildScatterChartEditSource();
    expect(chartXml(source).match(/<c:valAx>/g)).toHaveLength(2);
    expect(chartXml(source)).not.toContain("<c:catAx>");
    const chart = firstChart(source);
    const session = createEditorSession(source);
    const firstResult = session.updateScatterChartData(chart, {
      series: [
        { name: "Edited 1", xValues: [1, 2, 3], yValues: [3, 5, 8] },
        { name: "Edited 2", xValues: [10, 20], yValues: [2, 4] },
        { name: "Edited 3", xValues: [100], yValues: [6] },
      ],
    });
    expectApplied(firstResult);
    expect(chartXml(session.document)).toContain("Edited 1");
    expect(chartXml(session.document)).toContain("Sheet1!$A$2:$A$4");
    expect(chartXml(session.document)).toContain("Sheet1!$B$10");
    expect(chartXml(session.document).match(/<c:valAx>/g)).toHaveLength(2);
    expect(session.document.edits?.at(-1)?.kind).toBe("updateScatterChartData");
    expect(session.undoDepth).toBe(1);

    expect(chartXml(expectHistory(session.undo()))).not.toContain("Sheet1!$B$10");
    expect(chartXml(expectHistory(session.redo()))).toContain("Edited 3");

    const commandResult = session.apply({
      kind: "updateScatterChartData",
      handle: requireHandle(chart.handle),
      series: [{ name: "Command 1", xValues: [5, 10], yValues: [10, 20] }],
    });
    expectApplied(commandResult);
    expect(chartXml(session.document)).not.toContain("Sheet1!$B$6");
    expect(session.undoDepth).toBe(2);
    expect(chartXml(expectHistory(session.undo()))).toContain("Edited 3");
    expect(chartXml(expectHistory(session.redo()))).not.toContain("Sheet1!$B$6");
  });

  it("rejects invalid XY updates without changing document, selection, or history", async () => {
    const source = await buildScatterChartEditSource();
    const chart = firstChart(source);
    const session = createEditorSession(source);
    const handle = requireHandle(chart.handle);
    expect(session.selectShape(handle)).toMatchObject({ ok: true });
    const before = session.document;

    const result = session.apply({
      kind: "updateScatterChartData",
      handle,
      series: [{ name: "Invalid", xValues: [1, 2], yValues: [3] }],
    });

    expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    expect(result.ok ? undefined : result.message).toContain(
      "every series must have matching non-empty X and Y value counts",
    );
    expect(session.document).toBe(before);
    expect(session.selection).toEqual({ shapeHandle: handle });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });
});

describe("EditorSession bubble chart data commands", () => {
  it("supports convenience, command, undo, and redo while rejecting invalid XYZ atomically", async () => {
    const source = await buildBubbleChartEditSource();
    const chart = firstChart(source);
    const session = createEditorSession(source);
    const applied = session.updateBubbleChartData(chart, {
      series: [
        { name: "Edited 1", xValues: [1, 2], yValues: [3, 5], bubbleSizes: [8, 13] },
        { name: "Edited 2", xValues: [10], yValues: [20], bubbleSizes: [21] },
      ],
    });
    expectApplied(applied);
    expect(chartXml(session.document)).toContain("Sheet1!$C$2:$C$3");
    expect(session.document.edits?.at(-1)?.kind).toBe("updateBubbleChartData");
    expect(chartXml(expectHistory(session.undo()))).toContain("Original 1");
    expect(chartXml(expectHistory(session.redo()))).toContain("Edited 1");

    const command = session.apply({
      kind: "updateBubbleChartData",
      handle: requireHandle(chart.handle),
      series: [{ name: "Command", xValues: [5], yValues: [10], bubbleSizes: [15] }],
    });
    expectApplied(command);
    const before = session.document;
    const undoDepth = session.undoDepth;
    const invalid = session.apply({
      kind: "updateBubbleChartData",
      handle: requireHandle(chart.handle),
      series: [{ name: "Invalid", xValues: [1], yValues: [2], bubbleSizes: [] }],
    });
    expect(invalid).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(before);
    expect(session.undoDepth).toBe(undoDepth);
  });
});
