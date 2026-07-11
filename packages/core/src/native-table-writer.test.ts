import { addTable, asEmu, createPptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";

describe("native table writer rendering", () => {
  it("renders a table authored from scratch through the document reader path", async () => {
    const source = createPptx();
    const withTable = addTable(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(5486400),
      height: asEmu(1828800),
      columnWidths: [asEmu(2743200), asEmu(2743200)],
      rows: [
        {
          height: asEmu(914400),
          cells: [
            { text: "Header A", fill: "4472C4" },
            { text: "Header B", fill: "4472C4" },
          ],
        },
        {
          height: asEmu(914400),
          cells: [{ text: "Value A" }, { text: "Value B" }],
        },
      ],
    });

    const report = await renderPptxSourceModelToSvg(withTable, { skipSystemFonts: true });
    const svg = report.slides[0].svg;
    expect(svg).toContain("Header A");
    expect(svg).toContain("Value B");
    expect(svg).toContain("#4472c4");
  });
});
