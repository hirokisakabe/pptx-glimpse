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
    expect(slideXml).toContain('xml:space="preserve">Native ');
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

  it("writes cell margins, shared run formatting, line breaks, and schema-ordered properties", () => {
    const source = createPptx();
    const edited = addTable(source, source.slides[0].handle!, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(500),
      columnWidths: [asEmu(1000)],
      rows: [
        {
          height: asEmu(500),
          cells: [
            {
              runs: [
                {
                  text: "First\r\nSecond\nThird",
                  hyperlink: "https://example.com/formatted-table-run",
                  properties: {
                    bold: true,
                    strike: true,
                    baseline: "superscript",
                    highlight: { kind: "srgb", hex: "FFF2CC" },
                    underline: {
                      style: "dashLongHeavy",
                      color: { kind: "srgb", hex: "C00000" },
                    },
                  },
                },
              ],
              marginLeft: asEmu(100),
              marginRight: asEmu(200),
              marginTop: asEmu(300),
              marginBottom: asEmu(400),
              borders: {
                left: { dash: "lgDashDotDot" },
                right: { dash: "dashDot" },
                top: { dash: "sysDash" },
                bottom: { dash: "sysDot" },
              },
              fill: "D9EAF7",
            },
          ],
        },
      ],
    });

    const table = edited.slides[0].shapes.find((shape) => shape.kind === "table");
    const cell = table?.table.rows[0]?.cells[0];
    expect(cell).toMatchObject({
      marginLeft: 100,
      marginRight: 200,
      marginTop: 300,
      marginBottom: 400,
    });
    expect(cell?.textBody?.paragraphs[0]?.runs.map((run) => run.text)).toEqual([
      "First",
      "\n",
      "Second",
      "\n",
      "Third",
    ]);
    for (const run of cell?.textBody?.paragraphs[0]?.runs ?? []) {
      expect(run.properties).toMatchObject({
        bold: true,
        strikethrough: true,
        baseline: 30,
        underline: true,
        underlineStyle: "dashLongHeavy",
        underlineColor: { kind: "srgb", hex: "C00000" },
        highlight: { kind: "srgb", hex: "FFF2CC" },
      });
    }

    const output = writePptx(edited);
    const slideXml = decode(unzipSync(output)["ppt/slides/slide1.xml"]);
    expect(slideXml.match(/<a:br>/g)).toHaveLength(2);
    expect(slideXml).toContain('strike="sngStrike"');
    expect(slideXml).toContain('baseline="30000"');
    expect(slideXml).toContain('u="dashLongHeavy"');
    expect(slideXml).toContain('<a:uFill><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>');
    expect(slideXml).toContain('<a:highlight><a:srgbClr val="FFF2CC"/></a:highlight>');
    expect(slideXml).toContain('<a:hlinkClick r:id="rId2"/>');
    expect(slideXml).toContain('marL="100" marR="200" marT="300" marB="400"');
    expect(slideXml).toContain('<a:prstDash val="lgDashDotDot"/>');

    const tcPr = slideXml.slice(slideXml.indexOf("<a:tcPr"), slideXml.indexOf("</a:tcPr>"));
    expect(tcPr.indexOf("<a:lnL")).toBeLessThan(tcPr.indexOf("<a:lnR"));
    expect(tcPr.indexOf("<a:lnR")).toBeLessThan(tcPr.indexOf("<a:lnT"));
    expect(tcPr.indexOf("<a:lnT")).toBeLessThan(tcPr.indexOf("<a:lnB"));
    expect(tcPr.indexOf("<a:lnB")).toBeLessThan(tcPr.lastIndexOf("<a:solidFill"));

    const reread = readPptx(output);
    const rereadTable = reread.slides[0].shapes.find((shape) => shape.kind === "table");
    expect(rereadTable?.table.rows[0]?.cells[0]).toMatchObject({
      marginLeft: 100,
      marginRight: 200,
      marginTop: 300,
      marginBottom: 400,
      borders: { left: { dashStyle: "lgDashDotDot" } },
    });
  });

  it.each([
    "solid",
    "dash",
    "dashDot",
    "lgDash",
    "lgDashDot",
    "lgDashDotDot",
    "sysDash",
    "sysDot",
  ] as const)("supports the %s table border dash style", (dash) => {
    const source = createPptx();
    const edited = addTable(source, source.slides[0].handle!, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(100),
      height: asEmu(100),
      columnWidths: [asEmu(100)],
      rows: [{ height: asEmu(100), cells: [{ borders: { left: { dash } } }] }],
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const table = reread.slides[0].shapes.find((shape) => shape.kind === "table");
    expect(table?.table.rows[0]?.cells[0]?.borders?.left?.dashStyle).toBe(dash);
  });

  it("rejects invalid and ambiguous spans", () => {
    const source = createPptx();
    const base = {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(200),
      height: asEmu(100),
      columnWidths: [asEmu(100), asEmu(100)],
    } as const;
    expect(() =>
      addTable(source, source.slides[0].handle!, {
        ...base,
        rows: [{ height: asEmu(100), cells: [{ colspan: 1.5 }, {}] }],
      }),
    ).toThrow("spans must be positive integers");
    expect(() =>
      addTable(source, source.slides[0].handle!, {
        ...base,
        rows: [{ height: asEmu(100), cells: [{ colspan: 2 }, { text: "covered" }] }],
      }),
    ).toThrow("cells covered by a span must be empty placeholders");
  });
});
