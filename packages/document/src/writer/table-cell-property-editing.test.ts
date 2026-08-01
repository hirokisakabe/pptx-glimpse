import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addTable,
  asEmu,
  clearTableCellProperties,
  createPptx,
  type EditableTableCellProperties,
  readPptx,
  setTableCellProperties,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTable,
  writePptx,
} from "../index.js";
import { unsafeFixtureAssertion } from "../unsafe-type-assertion.js";

const decoder = new TextDecoder();
const encoder = new TextEncoder();

describe("writePptx - existing table cell property edits", () => {
  it("sets only the addressed fill, border, and margins while preserving cell content", () => {
    const source = readPptx(buildFixture());
    const table = firstTable(source);
    const address = { tableHandle: requireHandle(table.handle), rowIndex: 0, cellIndex: 0 };
    const edited = setTableCellProperties(source, address, {
      fill: { kind: "solid", color: { kind: "srgb", hex: "F4B183" } },
      borders: {
        top: {
          width: asEmu(25400),
          fill: { kind: "solid", color: { kind: "srgb", hex: "C00000" } },
        },
        right: { fill: { kind: "none" } },
      },
      marginLeft: asEmu(500),
      marginBottom: asEmu(600),
    });

    expect(source.edits).toBeUndefined();
    expect(firstTable(edited).table.rows[0].cells[0]).toMatchObject({
      fill: { kind: "solid", color: { kind: "srgb", hex: "F4B183" } },
      marginLeft: 500,
      marginRight: 200,
      marginTop: 300,
      marginBottom: 600,
      borders: {
        left: { width: 12700 },
        top: { width: 25400, fill: { kind: "solid" } },
        right: { fill: { kind: "none" } },
      },
    });

    const output = writePptx(edited);
    const reread = readPptx(output);
    const cell = firstTable(reread).table.rows[0].cells[0];
    const slideXml = decoder.decode(unzipSync(output)["ppt/slides/slide1.xml"]);
    expect(cell.textBody?.paragraphs[0].runs[0].text).toBe("Keep this text");
    expect(cell.fill).toEqual({ kind: "solid", color: { kind: "srgb", hex: "F4B183" } });
    expect(cell.borders?.top).toMatchObject({
      width: 25400,
      fill: { kind: "solid", color: { kind: "srgb", hex: "C00000" } },
    });
    expect(cell.borders?.right?.fill).toEqual({ kind: "none" });
    expect(cell).toMatchObject({
      marginLeft: 500,
      marginRight: 200,
      marginTop: 300,
      marginBottom: 600,
    });
    expect(firstTable(reread).table.rows[0].cells[1]).toMatchObject({ gridSpan: 2 });
    expect(firstTable(reread).table.rows[0].cells[2]).toMatchObject({ hMerge: true });
    expect(slideXml).toContain('uri="preserve-cell-sidecar"');
    expect(slideXml).toContain("<p:timing>");
    expect(decoder.decode(unzipSync(output)["customXml/item1.xml"])).toContain("preserve-part");
  });

  it("clears inline properties to inheritance without changing unrelated properties", () => {
    const source = readPptx(buildFixture());
    const table = firstTable(source);
    const edited = clearTableCellProperties(
      source,
      { tableHandle: requireHandle(table.handle), rowIndex: 0, cellIndex: 0 },
      ["fill", "borderLeft", "marginLeft", "marginTop"],
    );
    const output = writePptx(edited);
    const cell = firstTable(readPptx(output)).table.rows[0].cells[0];
    const slideXml = decoder.decode(unzipSync(output)["ppt/slides/slide1.xml"]);

    expect(cell.fill).toBeUndefined();
    expect(cell.borders?.left).toBeUndefined();
    expect(cell.marginLeft).toBeUndefined();
    expect(cell.marginTop).toBeUndefined();
    expect(cell.marginRight).toBe(asEmu(200));
    expect(cell.marginBottom).toBe(asEmu(400));
    expect(cell.textBody?.paragraphs[0].runs[0].text).toBe("Keep this text");
    expect(slideXml).toContain('uri="preserve-cell-sidecar"');
  });

  it("returns no-ops unchanged and rejects invalid addresses and unsupported styles atomically", () => {
    const source = readPptx(buildFixture());
    const table = firstTable(source);
    const address = { tableHandle: requireHandle(table.handle), rowIndex: 0, cellIndex: 0 };
    const noOp = setTableCellProperties(source, address, {
      fill: { kind: "solid", color: { kind: "srgb", hex: "D9EAF7" } },
      marginLeft: asEmu(100),
    });

    expect(noOp).toBe(source);
    expect(source.edits).toBeUndefined();
    expect(() =>
      setTableCellProperties(source, { ...address, rowIndex: 99 }, { marginLeft: asEmu(1) }),
    ).toThrow("table cell address was not found");
    expect(() =>
      setTableCellProperties(
        source,
        address,
        unsafeFixtureAssertion<EditableTableCellProperties>({
          fill: { kind: "gradient", stops: [] },
        }),
      ),
    ).toThrow("supports only solid and none fills");
    expect(() =>
      setTableCellProperties(
        source,
        address,
        unsafeFixtureAssertion<EditableTableCellProperties>({
          borders: { left: { width: asEmu(12700), dashStyle: "dash" } },
        }),
      ),
    ).toThrow("is not a supported border style");
    expect(source.edits).toBeUndefined();
    expect(firstTable(source).table.rows[0].cells[0].fill).toMatchObject({
      kind: "solid",
      color: { hex: "D9EAF7" },
    });
  });
});

function buildFixture(): Uint8Array {
  const source = createPptx();
  const authored = addTable(source, source.slides[0].handle!, {
    offsetX: asEmu(457200),
    offsetY: asEmu(457200),
    width: asEmu(8229600),
    height: asEmu(1828800),
    columnWidths: [asEmu(2743200), asEmu(2743200), asEmu(2743200)],
    tableStyleId: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    rows: [
      {
        height: asEmu(914400),
        cells: [
          {
            text: "Keep this text",
            fill: "D9EAF7",
            marginLeft: asEmu(100),
            marginRight: asEmu(200),
            marginTop: asEmu(300),
            marginBottom: asEmu(400),
            borders: { left: { width: asEmu(12700), color: "4472C4" } },
          },
          { text: "Merged", colspan: 2 },
          {},
        ],
      },
      {
        height: asEmu(914400),
        cells: [{ text: "Sibling A" }, { text: "Sibling B" }, { text: "Sibling C" }],
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

function firstTable(source: {
  readonly slides: readonly { readonly shapes: readonly SourceShapeNode[] }[];
}): SourceTable {
  const table = findTable(source.slides[0].shapes);
  if (table === undefined) throw new Error("test fixture table is missing");
  return table;
}

function findTable(nodes: readonly SourceShapeNode[]): SourceTable | undefined {
  for (const node of nodes) {
    if (node.kind === "table") return node;
    if (node.kind === "group") {
      const nested = findTable(node.children);
      if (nested !== undefined) return nested;
    }
  }
  return undefined;
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("test fixture handle is missing");
  return handle;
}
