import { createComputedView, readPptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createPptxEditorSession } from "./index.js";
import { buildScatterChartFixture } from "./pptx-editor-session.test-helpers.js";

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
    const chartData = createComputedView(readPptx(saved.pptx)).slides[0]?.elements.find(
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
