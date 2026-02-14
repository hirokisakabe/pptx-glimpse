import { describe, it, expect } from "vitest";
import { parseTable } from "./table-parser.js";
import { ColorResolver } from "../color/color-resolver.js";

function createColorResolver() {
  return new ColorResolver(
    {
      dk1: "#000000",
      lt1: "#FFFFFF",
      dk2: "#44546A",
      lt2: "#E7E6E6",
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
      hlink: "#0563C1",
      folHlink: "#954F72",
    },
    {
      bg1: "lt1",
      tx1: "dk1",
      bg2: "lt2",
      tx2: "dk2",
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hlink: "hlink",
      folHlink: "folHlink",
    },
  );
}

describe("parseTable", () => {
  it("parses a basic 2x2 table", () => {
    const tblNode = {
      tblGrid: {
        gridCol: [{ "@_w": "914400" }, { "@_w": "914400" }],
      },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              txBody: {
                bodyPr: {},
                p: [{ r: [{ t: "A1", rPr: {} }] }],
              },
              tcPr: {},
            },
            {
              txBody: {
                bodyPr: {},
                p: [{ r: [{ t: "B1", rPr: {} }] }],
              },
              tcPr: {},
            },
          ],
        },
        {
          "@_h": "457200",
          tc: [
            {
              txBody: {
                bodyPr: {},
                p: [{ r: [{ t: "A2", rPr: {} }] }],
              },
              tcPr: {},
            },
            {
              txBody: {
                bodyPr: {},
                p: [{ r: [{ t: "B2", rPr: {} }] }],
              },
              tcPr: {},
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    expect(result).not.toBeNull();
    expect(result!.columns).toHaveLength(2);
    expect(result!.columns[0].width).toBe(914400);
    expect(result!.rows).toHaveLength(2);
    expect(result!.rows[0].height).toBe(457200);
    expect(result!.rows[0].cells).toHaveLength(2);
    expect(result!.rows[0].cells[0].textBody).not.toBeNull();
    expect(result!.rows[0].cells[0].textBody!.paragraphs[0].runs[0].text).toBe("A1");
  });

  it("returns null when tblNode is null", () => {
    const result = parseTable(null, createColorResolver());
    expect(result).toBeNull();
  });

  it("returns null when tblGrid has no columns", () => {
    const tblNode = { tblGrid: { gridCol: [] }, tr: [] };
    const result = parseTable(tblNode, createColorResolver());
    expect(result).toBeNull();
  });

  it("parses cell fill", () => {
    const tblNode = {
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              txBody: null,
              tcPr: {
                solidFill: { srgbClr: { "@_val": "FF0000" } },
              },
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    expect(result!.rows[0].cells[0].fill).not.toBeNull();
    expect(result!.rows[0].cells[0].fill!.type).toBe("solid");
  });

  it("parses cell borders", () => {
    const tblNode = {
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              txBody: null,
              tcPr: {
                lnT: { "@_w": "12700", solidFill: { srgbClr: { "@_val": "000000" } } },
                lnB: { "@_w": "12700", solidFill: { srgbClr: { "@_val": "000000" } } },
                lnL: { "@_w": "12700", solidFill: { srgbClr: { "@_val": "000000" } } },
                lnR: { "@_w": "12700", solidFill: { srgbClr: { "@_val": "000000" } } },
              },
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    const borders = result!.rows[0].cells[0].borders;
    expect(borders).not.toBeNull();
    expect(borders!.top).not.toBeNull();
    expect(borders!.bottom).not.toBeNull();
    expect(borders!.left).not.toBeNull();
    expect(borders!.right).not.toBeNull();
  });

  it("parses cell merge attributes", () => {
    const tblNode = {
      tblGrid: { gridCol: [{ "@_w": "914400" }, { "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              "@_gridSpan": "2",
              txBody: null,
              tcPr: {},
            },
            {
              txBody: null,
              tcPr: { "@_hMerge": "1" },
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    expect(result!.rows[0].cells[0].gridSpan).toBe(2);
    expect(result!.rows[0].cells[0].hMerge).toBe(false);
    expect(result!.rows[0].cells[1].hMerge).toBe(true);
  });

  it("parses vertical merge attributes", () => {
    const tblNode = {
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              "@_rowSpan": "2",
              txBody: null,
              tcPr: {},
            },
          ],
        },
        {
          "@_h": "457200",
          tc: [
            {
              txBody: null,
              tcPr: { "@_vMerge": "1" },
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    expect(result!.rows[0].cells[0].rowSpan).toBe(2);
    expect(result!.rows[1].cells[0].vMerge).toBe(true);
  });

  it("handles cells with no tcPr", () => {
    const tblNode = {
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [{ txBody: null }],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    expect(result!.rows[0].cells[0].fill).toBeNull();
    expect(result!.rows[0].cells[0].borders).toBeNull();
  });

  it("applies default borders when tableStyleId is present and no inline borders", () => {
    const tblNode = {
      tblPr: {
        tableStyleId: "{D198CDD0-A2AB-4776-9A92-BCB6353A44E2}",
      },
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [{ txBody: null, tcPr: {} }],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    const borders = result!.rows[0].cells[0].borders;
    expect(borders).not.toBeNull();
    expect(borders!.top).not.toBeNull();
    expect(borders!.bottom).not.toBeNull();
    expect(borders!.left).not.toBeNull();
    expect(borders!.right).not.toBeNull();
    expect(borders!.top!.fill).toEqual({
      type: "solid",
      color: { hex: "#000000", alpha: 1 },
    });
  });

  it("does not apply default borders when no tableStyleId", () => {
    const tblNode = {
      tblPr: {},
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [{ txBody: null, tcPr: {} }],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    const borders = result!.rows[0].cells[0].borders;
    expect(borders).toBeNull();
  });

  it("inline borders take precedence over table-style defaults", () => {
    const tblNode = {
      tblPr: {
        tableStyleId: "{D198CDD0-A2AB-4776-9A92-BCB6353A44E2}",
      },
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              txBody: null,
              tcPr: {
                lnT: { "@_w": "25400", solidFill: { srgbClr: { "@_val": "FF0000" } } },
                lnB: { "@_w": "25400", solidFill: { srgbClr: { "@_val": "FF0000" } } },
                lnL: { "@_w": "25400", solidFill: { srgbClr: { "@_val": "FF0000" } } },
                lnR: { "@_w": "25400", solidFill: { srgbClr: { "@_val": "FF0000" } } },
              },
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    const borders = result!.rows[0].cells[0].borders;
    expect(borders).not.toBeNull();
    expect(borders!.top!.width).toBe(25400);
    expect(borders!.top!.fill).toEqual({
      type: "solid",
      color: { hex: "#FF0000", alpha: 1 },
    });
  });

  it("applies default black fill to border with width but no color", () => {
    const tblNode = {
      tblGrid: { gridCol: [{ "@_w": "914400" }] },
      tr: [
        {
          "@_h": "457200",
          tc: [
            {
              txBody: null,
              tcPr: {
                lnT: { "@_w": "12700" },
              },
            },
          ],
        },
      ],
    };

    const result = parseTable(tblNode, createColorResolver());
    const borders = result!.rows[0].cells[0].borders;
    expect(borders).not.toBeNull();
    expect(borders!.top).not.toBeNull();
    expect(borders!.top!.fill).toEqual({
      type: "solid",
      color: { hex: "#000000", alpha: 1 },
    });
  });
});
