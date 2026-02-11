import { describe, it, expect } from "vitest";
import { renderTable } from "./table-renderer.js";
import type { TableElement } from "../model/table.js";

function createTableElement(overrides?: Partial<TableElement["table"]>): TableElement {
  return {
    type: "table",
    transform: {
      offsetX: 914400,
      offsetY: 914400,
      extentWidth: 1828800,
      extentHeight: 914400,
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    table: {
      columns: [{ width: 914400 }, { width: 914400 }],
      rows: [
        {
          height: 457200,
          cells: [
            {
              textBody: null,
              fill: { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
              borders: null,
              gridSpan: 1,
              rowSpan: 1,
              hMerge: false,
              vMerge: false,
            },
            {
              textBody: null,
              fill: { type: "solid", color: { hex: "#00FF00", alpha: 1 } },
              borders: null,
              gridSpan: 1,
              rowSpan: 1,
              hMerge: false,
              vMerge: false,
            },
          ],
        },
        {
          height: 457200,
          cells: [
            {
              textBody: null,
              fill: { type: "solid", color: { hex: "#0000FF", alpha: 1 } },
              borders: null,
              gridSpan: 1,
              rowSpan: 1,
              hMerge: false,
              vMerge: false,
            },
            {
              textBody: null,
              fill: { type: "solid", color: { hex: "#FFFF00", alpha: 1 } },
              borders: null,
              gridSpan: 1,
              rowSpan: 1,
              hMerge: false,
              vMerge: false,
            },
          ],
        },
      ],
      ...overrides,
    },
  };
}

describe("renderTable", () => {
  it("renders a basic 2x2 table with cell backgrounds", () => {
    const element = createTableElement();
    const defs: string[] = [];
    const svg = renderTable(element, defs);

    expect(svg).toContain("<g transform=");
    // 4 cells, each with a rect
    const rects = svg.match(/<rect /g);
    expect(rects).toHaveLength(4);
    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain('fill="#00FF00"');
    expect(svg).toContain('fill="#0000FF"');
    expect(svg).toContain('fill="#FFFF00"');
  });

  it("renders cell borders as lines", () => {
    const element: TableElement = {
      type: "table",
      transform: {
        offsetX: 0,
        offsetY: 0,
        extentWidth: 914400,
        extentHeight: 457200,
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      table: {
        columns: [{ width: 914400 }],
        rows: [
          {
            height: 457200,
            cells: [
              {
                textBody: null,
                fill: null,
                borders: {
                  top: {
                    width: 12700,
                    fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
                    dashStyle: "solid",
                  },
                  bottom: {
                    width: 12700,
                    fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
                    dashStyle: "solid",
                  },
                  left: {
                    width: 12700,
                    fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
                    dashStyle: "solid",
                  },
                  right: {
                    width: 12700,
                    fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
                    dashStyle: "solid",
                  },
                },
                gridSpan: 1,
                rowSpan: 1,
                hMerge: false,
                vMerge: false,
              },
            ],
          },
        ],
      },
    };

    const defs: string[] = [];
    const svg = renderTable(element, defs);
    const lines = svg.match(/<line /g);
    expect(lines).toHaveLength(4);
  });

  it("skips hMerge and vMerge cells", () => {
    const element: TableElement = {
      type: "table",
      transform: {
        offsetX: 0,
        offsetY: 0,
        extentWidth: 1828800,
        extentHeight: 457200,
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      table: {
        columns: [{ width: 914400 }, { width: 914400 }],
        rows: [
          {
            height: 457200,
            cells: [
              {
                textBody: null,
                fill: { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
                borders: null,
                gridSpan: 2,
                rowSpan: 1,
                hMerge: false,
                vMerge: false,
              },
              {
                textBody: null,
                fill: null,
                borders: null,
                gridSpan: 1,
                rowSpan: 1,
                hMerge: true,
                vMerge: false,
              },
            ],
          },
        ],
      },
    };

    const defs: string[] = [];
    const svg = renderTable(element, defs);
    // Only 1 rect for the merged cell, hMerge cell is skipped
    const rects = svg.match(/<rect /g);
    expect(rects).toHaveLength(1);
  });

  it("renders cell text inside a translated group", () => {
    const element: TableElement = {
      type: "table",
      transform: {
        offsetX: 0,
        offsetY: 0,
        extentWidth: 914400,
        extentHeight: 457200,
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      table: {
        columns: [{ width: 914400 }],
        rows: [
          {
            height: 457200,
            cells: [
              {
                textBody: {
                  paragraphs: [
                    {
                      runs: [
                        {
                          text: "Hello",
                          properties: {
                            fontSize: 12,
                            fontFamily: null,
                            fontFamilyEa: null,
                            bold: false,
                            italic: false,
                            underline: false,
                            strikethrough: false,
                            color: null,
                            baseline: 0,
                          },
                        },
                      ],
                      properties: {
                        alignment: "l",
                        lineSpacing: null,
                        spaceBefore: 0,
                        spaceAfter: 0,
                        level: 0,
                      },
                    },
                  ],
                  bodyProperties: {
                    anchor: "t",
                    marginLeft: 91440,
                    marginRight: 91440,
                    marginTop: 45720,
                    marginBottom: 45720,
                    wrap: "square",
                    autoFit: "noAutofit",
                    fontScale: 1,
                    lnSpcReduction: 0,
                  },
                },
                fill: null,
                borders: null,
                gridSpan: 1,
                rowSpan: 1,
                hMerge: false,
                vMerge: false,
              },
            ],
          },
        ],
      },
    };

    const defs: string[] = [];
    const svg = renderTable(element, defs);
    expect(svg).toContain("<text");
    expect(svg).toContain("Hello");
  });
});
