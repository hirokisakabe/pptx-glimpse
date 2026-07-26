import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addTable,
  addTextBox,
  asEmu,
  asPt,
  clearParagraphProperties,
  clearTextRunProperties,
  createPptx,
  findParagraphBySourceHandle,
  findTextRunBySourceHandle,
  type PptxSourceModel,
  readPptx,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setTextRunProperties,
  type SourceHandle,
  type SourceTable,
  writePptx,
} from "../index.js";

const decoder = new TextDecoder();
const encoder = new TextEncoder();

describe("writePptx - existing table cell text edits", () => {
  it("reuses public text operations and preserves the surrounding table structure", () => {
    const input = buildExistingTableFixture();
    const source = readPptx(input);
    const table = firstTable(source);
    const runPlainText = table.table.rows[0].cells[0].textBody!.paragraphs[0].runs[0];
    const paragraphPlainText = table.table.rows[0].cells[1].textBody!.paragraphs[0];
    const runProperties = table.table.rows[1].cells[0].textBody!.paragraphs[0].runs[0];
    const paragraphProperties = table.table.rows[1].cells[1].textBody!.paragraphs[0];

    expect(runPlainText.handle).toMatchObject({
      nodeId: "text:table:1:row:0:cell:0:p:0:r:0",
    });
    expect(paragraphPlainText.handle).toMatchObject({
      nodeId: "text:table:1:row:0:cell:1:p:0",
    });
    expect(findTextRunBySourceHandle(source, requireHandle(runPlainText.handle))).toBe(
      runPlainText,
    );
    expect(findParagraphBySourceHandle(source, requireHandle(paragraphPlainText.handle))).toBe(
      paragraphPlainText,
    );

    let edited = replaceTextRunPlainText(
      source,
      requireHandle(runPlainText.handle),
      "Edited table run",
    );
    edited = replaceParagraphPlainText(
      edited,
      requireHandle(paragraphPlainText.handle),
      "Edited table paragraph",
    );
    edited = setTextRunProperties(edited, requireHandle(runProperties.handle), {
      bold: true,
      italic: true,
      fontSize: asPt(22),
      color: { kind: "srgb", hex: "AA0000" },
      typeface: "Aptos",
    });
    edited = clearTextRunProperties(edited, requireHandle(runProperties.handle), ["italic"]);
    edited = setParagraphProperties(edited, requireHandle(paragraphProperties.handle), {
      align: "right",
      level: 1,
      bullet: { type: "char", char: "•" },
    });
    edited = clearParagraphProperties(edited, requireHandle(paragraphProperties.handle), ["level"]);

    const output = writePptx(edited);
    const reread = readPptx(output);
    const rereadTable = firstTable(reread);
    const slideXml = decoder.decode(unzipSync(output)["ppt/slides/slide1.xml"]);

    expect(rereadTable.table.rows[0].cells[0].textBody?.paragraphs[0].runs[0].text).toBe(
      "Edited table run",
    );
    expect(rereadTable.table.rows[0].cells[1].textBody?.paragraphs[0].runs).toMatchObject([
      { text: "Edited table paragraph" },
    ]);
    expect(
      rereadTable.table.rows[1].cells[0].textBody?.paragraphs[0].runs[0].properties,
    ).toMatchObject({
      bold: true,
      fontSize: 22,
      color: { kind: "srgb", hex: "AA0000" },
      typeface: "Aptos",
    });
    expect(
      rereadTable.table.rows[1].cells[0].textBody?.paragraphs[0].runs[0].properties?.italic,
    ).toBeUndefined();
    expect(rereadTable.table.rows[1].cells[1].textBody?.paragraphs[0].properties).toMatchObject({
      align: "right",
      bullet: { type: "char", char: "•" },
    });
    expect(
      rereadTable.table.rows[1].cells[1].textBody?.paragraphs[0].properties?.level,
    ).toBeUndefined();

    expect(rereadTable.table.tableStyleId).toBe("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
    expect(rereadTable.table.rows[0].cells[0]).toMatchObject({
      fill: { kind: "solid", color: { kind: "srgb", hex: "D9EAF7" } },
      marginLeft: 100,
      marginRight: 200,
      marginTop: 300,
      marginBottom: 400,
    });
    expect(rereadTable.table.rows[0].cells[0].borders?.left).toMatchObject({
      width: 12700,
      fill: { kind: "solid", color: { kind: "srgb", hex: "4472C4" } },
    });
    expect(rereadTable.table.rows[0].cells[2]).toMatchObject({ gridSpan: 2 });
    expect(rereadTable.table.rows[0].cells[3]).toMatchObject({ hMerge: true });
    expect(rereadTable.table.rows[1].cells[2].textBody?.paragraphs[0].runs[0].text).toBe(
      "Unedited cell",
    );
    expect(slideXml).toContain('uri="preserve-cell-sidecar"');
    expect(slideXml).toContain("<p:timing>");
    expect(decoder.decode(unzipSync(output)["customXml/item1.xml"])).toContain("preserve-part");
  });

  it("rejects a foreign table text handle without changing the source", () => {
    const source = readPptx(buildExistingTableFixture());
    const other = readPptx(buildExistingTableFixture("Other", true));
    const foreignRun = firstTable(other).table.rows[0].cells[0].textBody!.paragraphs[0].runs[0];
    expect(() =>
      replaceTextRunPlainText(source, requireHandle(foreignRun.handle), "Rejected"),
    ).toThrow("text run handle was not found");
    expect(source.edits).toBeUndefined();
    expect(firstTable(source).table.rows[0].cells[0].textBody?.paragraphs[0].runs[0].text).toBe(
      "Original table run",
    );
  });

  it("rejects conflicting run and paragraph edits within the same table cell", () => {
    const source = readPptx(buildExistingTableFixture());
    const paragraph = firstTable(source).table.rows[0].cells[0].textBody!.paragraphs[0];
    const runHandle = requireHandle(paragraph.runs[0].handle);
    const paragraphHandle = requireHandle(paragraph.handle);
    const textConflict = replaceParagraphPlainText(
      replaceTextRunPlainText(source, runHandle, "Run edit"),
      paragraphHandle,
      "Paragraph edit",
    );
    const propertyConflict = replaceParagraphPlainText(
      setTextRunProperties(source, runHandle, { bold: true }),
      paragraphHandle,
      "Paragraph edit",
    );

    expect(() => writePptx(textConflict)).toThrow(/conflicting text run and paragraph edits/);
    expect(() => writePptx(propertyConflict)).toThrow(
      /conflicting text run properties and paragraph edits/,
    );
  });
});

function buildExistingTableFixture(
  firstCellText = "Original table run",
  withLeadingShape = false,
): Uint8Array {
  const source = createPptx();
  const withOptionalShape = withLeadingShape
    ? addTextBox(source, source.slides[0].handle!, {
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(100),
        height: asEmu(100),
        text: "Leading shape",
      })
    : source;
  const authored = addTable(withOptionalShape, withOptionalShape.slides[0].handle!, {
    offsetX: asEmu(457200),
    offsetY: asEmu(457200),
    width: asEmu(8229600),
    height: asEmu(1828800),
    columnWidths: [asEmu(2057400), asEmu(2057400), asEmu(2057400), asEmu(2057400)],
    tableStyleId: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    rows: [
      {
        height: asEmu(914400),
        cells: [
          {
            text: firstCellText,
            fill: "D9EAF7",
            marginLeft: asEmu(100),
            marginRight: asEmu(200),
            marginTop: asEmu(300),
            marginBottom: asEmu(400),
            borders: { left: { width: asEmu(12700), color: "4472C4" } },
          },
          { text: "Original table paragraph" },
          { text: "Merged", colspan: 2 },
          {},
        ],
      },
      {
        height: asEmu(914400),
        cells: [
          { text: "Run properties" },
          { text: "Paragraph properties" },
          { text: "Unedited cell" },
          { text: "Unedited sibling" },
        ],
      },
    ],
  });
  const files = unzipSync(writePptx(authored));
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = decoder
    .decode(files[slidePath])
    .replace("</a:tcPr>", '<a:extLst><a:ext uri="preserve-cell-sidecar"/></a:extLst></a:tcPr>')
    .replace("</p:sld>", "<p:timing><p:tnLst><p:par/></p:tnLst></p:timing></p:sld>");
  return zipSync({
    ...files,
    [slidePath]: encoder.encode(slideXml),
    "customXml/item1.xml": encoder.encode("<root>preserve-part</root>"),
  });
}

function firstTable(source: PptxSourceModel): SourceTable {
  const table = source.slides[0].shapes.find((shape) => shape.kind === "table");
  if (table === undefined) throw new Error("test fixture table is missing");
  return table;
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("test fixture handle is missing");
  return handle;
}
