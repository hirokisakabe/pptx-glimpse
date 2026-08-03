import { createComputedView, readPptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createPptxEditorSession } from "./index.js";
import {
  buildBubbleChartFixture,
  buildScatterChartFixture,
} from "./pptx-editor-session.test-helpers.js";

describe("PptxEditorSession - scatter chart data", () => {
  it("applies, renders, saves, and restores the scatter command through history", async () => {
    const editor = await createPptxEditorSession(await buildScatterChartFixture(), {
      skipSystemFonts: true,
    });
    const chart = editor.document.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("scatter fixture has no chart handle");
    const beforeSvg = editor.slides[0]?.svg;

    const applied = await editor.apply({
      kind: "updateScatterChartData",
      handle: chart.handle,
      series: [
        { name: "Edited 1", xValues: [1, 2, 3], yValues: [8, 3, 6] },
        { name: "Edited 2", xValues: [10, 20], yValues: [5, 7] },
      ],
    });

    expect(applied.history).toMatchObject({ canUndo: true, undoDepth: 1, redoDepth: 0 });
    expect(applied.slides[0]?.svg).not.toBe(beforeSvg);
    const saved = editor.save();
    const savedDocument = readPptx(saved.pptx);
    const savedChartPart = savedDocument.packageGraph.rawParts?.find(
      (part) => part.partPath === "ppt/charts/chart1.xml",
    );
    if (savedChartPart?.kind !== "binary") throw new Error("saved scatter chart XML is missing");
    const savedChartXml = new TextDecoder().decode(savedChartPart.bytes);
    expect(savedChartXml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(savedChartXml).not.toContain("<c:catAx>");
    const chartData = createComputedView(savedDocument).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(chartData?.kind === "chart" ? chartData.chartData?.series : undefined).toMatchObject([
      { name: "Edited 1", xValues: [1, 2, 3], values: [8, 3, 6] },
      { name: "Edited 2", xValues: [10, 20], values: [5, 7] },
    ]);

    expect((await editor.undo()).history).toMatchObject({ undoDepth: 0, redoDepth: 1 });
    expect((await editor.redo()).history).toMatchObject({ undoDepth: 1, redoDepth: 0 });
  });
});

describe("PptxEditorSession - bubble chart data", () => {
  it("applies, renders, saves, and restores the bubble command through history", async () => {
    const editor = await createPptxEditorSession(await buildBubbleChartFixture(), {
      skipSystemFonts: true,
    });
    const chart = editor.document.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("bubble fixture has no chart handle");
    const beforeSvg = editor.slides[0]?.svg;

    const applied = await editor.apply({
      kind: "updateBubbleChartData",
      handle: chart.handle,
      series: [
        { name: "Edited 1", xValues: [1, 2, 3], yValues: [8, 3, 6], bubbleSizes: [4, 9, 16] },
        { name: "Edited 2", xValues: [10], yValues: [5], bubbleSizes: [25] },
      ],
    });

    expect(applied.history).toMatchObject({ canUndo: true, undoDepth: 1, redoDepth: 0 });
    expect(applied.slides[0]?.svg).not.toBe(beforeSvg);
    const savedDocument = readPptx(editor.save().pptx);
    const chartData = createComputedView(savedDocument).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(chartData?.kind === "chart" ? chartData.chartData : undefined).toMatchObject({
      chartType: "bubble",
      series: [
        { name: "Edited 1", xValues: [1, 2, 3], values: [8, 3, 6], bubbleSizes: [4, 9, 16] },
        { name: "Edited 2", xValues: [10], values: [5], bubbleSizes: [25] },
      ],
    });

    expect((await editor.undo()).history).toMatchObject({ undoDepth: 0, redoDepth: 1 });
    expect((await editor.redo()).history).toMatchObject({ undoDepth: 1, redoDepth: 0 });
  });
});
