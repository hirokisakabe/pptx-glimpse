import {
  addTable,
  asEmu,
  createPptx,
  readPptx,
  setTableCellProperties,
  writePptx,
} from "@pptx-glimpse/document";
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

  it("writes, rereads, and renders edited existing cell properties", async () => {
    const source = createPptx();
    const authored = addTable(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(5486400),
      height: asEmu(914400),
      columnWidths: [asEmu(2743200), asEmu(2743200)],
      rows: [{ height: asEmu(914400), cells: [{ text: "Edited cell" }, { text: "Sibling" }] }],
    });
    const existing = readPptx(writePptx(authored));
    const table = existing.slides[0].shapes.find((shape) => shape.kind === "table");
    if (table?.handle === undefined) throw new Error("existing table should have a handle");
    const edited = setTableCellProperties(
      existing,
      { tableHandle: table.handle, rowIndex: 0, cellIndex: 0 },
      {
        fill: { kind: "solid", color: { kind: "srgb", hex: "F4B183" } },
        borders: {
          bottom: {
            width: asEmu(25400),
            fill: { kind: "solid", color: { kind: "srgb", hex: "C00000" } },
          },
        },
        marginLeft: asEmu(457200),
      },
    );
    const reread = readPptx(writePptx(edited));
    const report = await renderPptxSourceModelToSvg(reread, { skipSystemFonts: true });
    const svg = report.slides[0].svg;

    expect(svg).toContain("Edited cell");
    expect(svg).toContain("#f4b183");
    expect(svg).toContain("#c00000");
  });
});
