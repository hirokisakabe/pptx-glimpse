import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import { readPptx } from "../reader/read-pptx.js";
import { writePptx } from "../writer/write-pptx.js";
import { addTable } from "./table-authoring.js";
import { asEmu, asPt } from "./units.js";

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe("addTable", () => {
  it("writes a native table with rich cells, merges, hyperlinks, and consistent ids", () => {
    const source = createPptx();
    const edited = addTable(source, source.slides[0].handle!, {
      offsetX: asEmu(457200),
      offsetY: asEmu(457200),
      width: asEmu(8229600),
      height: asEmu(2743200),
      columnWidths: [asEmu(2743200), asEmu(2743200), asEmu(2743200)],
      rows: [
        {
          height: asEmu(914400),
          cells: [
            {
              runs: [
                { text: "Native ", properties: { bold: true, color: "FFFFFF" } },
                {
                  text: "Table",
                  properties: { italic: true, fontSize: asPt(18), fontFace: "Aptos" },
                  hyperlink: "https://example.com/table",
                },
              ],
              fill: "4472C4",
              align: "center",
              verticalAlign: "middle",
              colspan: 2,
              borders: { bottom: { width: asEmu(12700), color: "FFFFFF" } },
            },
            {},
            { text: "rowspan", rowspan: 2 },
          ],
        },
        {
          height: asEmu(914400),
          cells: [{ text: "A" }, { text: "B" }, {}],
        },
        {
          height: asEmu(914400),
          cells: [{ text: "C" }, { text: "D" }, { text: "E" }],
        },
      ],
    });

    const output = writePptx(edited);
    const files = unzipSync(output);
    const slideXml = decode(files["ppt/slides/slide1.xml"]);
    const relsXml = decode(files["ppt/slides/_rels/slide1.xml.rels"]);

    expect(slideXml).toContain("<p:graphicFrame>");
    expect(slideXml).toContain("<a:tbl>");
    expect(slideXml).toContain('gridSpan="2"');
    expect(slideXml).toContain('hMerge="1"');
    expect(slideXml).toContain('rowSpan="2"');
    expect(slideXml).toContain('vMerge="1"');
    expect(slideXml).toContain('<a:hlinkClick r:id="rId2"');
    expect(relsXml).toContain('Id="rId2"');
    expect(relsXml).toContain('Target="https://example.com/table"');
    expect(relsXml).toContain('TargetMode="External"');

    const reread = readPptx(output);
    const table = reread.slides[0].shapes.find((shape) => shape.kind === "table");
    expect(table?.nodeId).toBe("2");
    expect(table?.table.columns).toHaveLength(3);
    expect(table?.table.rows).toHaveLength(3);
    expect(table?.table.rows[0].cells[0]).toMatchObject({ gridSpan: 2 });
    expect(table?.table.rows[0].cells[1]).toMatchObject({ hMerge: true });
    expect(table?.table.rows[1].cells[2]).toMatchObject({ vMerge: true });
  });
});
