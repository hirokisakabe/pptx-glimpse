import { describe, expect, it } from "vitest";

import type {
  CleanDocSource,
  ComputedConnectorElement,
  ComputedElement,
  ComputedGroupElement,
  ComputedShapeElement,
  ComputedSlide,
  ComputedTableElement,
  SourceRunProperties,
  SourceShape,
  SourceTable,
  SourceTextBody,
} from "../experimental.js";
import {
  asEmu,
  asOoxmlPercent,
  asPartPath,
  asPt,
  asRelationshipId,
  createComputedView,
} from "../experimental.js";

describe("createComputedView", () => {
  it("slide size / order / relationships を computed view に反映する", () => {
    const source = buildSource();
    const computed = createComputedView(source);

    expect(computed.slideSize).toEqual({ width: 9144000, height: 5143500 });
    expect(computed.slides.map((slide) => slide.partPath)).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);

    const slide = getSlide(computed.slides, 0);
    expect(slide.slideNumber).toBe(1);
    expect(slide.slideSize).toEqual({ width: 9144000, height: 5143500 });
    const imageRelationship = slide.relationships.find((rel) => rel.id === "rIdImage");
    expect(imageRelationship?.target).toBe("ppt/media/image1.png");
    expect(imageRelationship?.targetPartPath).toBe("ppt/media/image1.png");
    expect(imageRelationship?.media?.contentType).toBe("image/png");
    const externalRelationship = slide.relationships.find((rel) => rel.id === "rIdExternal");
    expect(externalRelationship?.target).toBe("https://example.com/");
    expect(externalRelationship?.targetMode).toBe("External");

    const image = slide.elements.find((element) => element.kind === "image");
    expect(image?.media?.bytes).toEqual(new Uint8Array([1, 2, 3]));
  });

  it("chart / SmartArt の relationship と source XML part を解決する", () => {
    const source = buildSourceWithChartAndSmartArt();
    const slide = getSlide(createComputedView(source).slides, 0);
    const chart = slide.elements.find((element) => element.kind === "chart");
    const smartArt = slide.elements.find((element) => element.kind === "smartArt");

    expect(chart?.kind).toBe("chart");
    if (chart?.kind !== "chart") throw new Error("chart element not found");
    expect(chart.relationship?.id).toBe("rIdChart");
    expect(chart.relationship?.targetPartPath).toBe("ppt/charts/chart1.xml");
    expect(chart.chartXml).toContain("<c:barChart");

    expect(smartArt?.kind).toBe("smartArt");
    if (smartArt?.kind !== "smartArt") throw new Error("SmartArt element not found");
    expect(smartArt.dataRelationship?.id).toBe("rIdDiagramData");
    expect(smartArt.dataRelationship?.targetPartPath).toBe("ppt/diagrams/data1.xml");
    expect(smartArt.drawingRelationship?.targetPartPath).toBe("ppt/diagrams/drawing1.xml");
    expect(smartArt.drawingPartPath).toBe("ppt/diagrams/drawing1.xml");
    expect(smartArt.drawingXml).toContain("<dsp:drawing");
    expect(smartArt.drawingRelationships).toHaveLength(0);
  });

  it("Strict OOXML の chart relationship type でも chart XML part を解決する", () => {
    const source = buildSourceWithChartAndSmartArt({
      chartRelationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
    });
    const slide = getSlide(createComputedView(source).slides, 0);
    const chart = slide.elements.find((element) => element.kind === "chart");

    expect(chart?.kind).toBe("chart");
    if (chart?.kind !== "chart") throw new Error("chart element not found");
    expect(chart.relationship?.type).toBe(
      "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
    );
    expect(chart.chartXml).toContain("<c:barChart");
  });

  it("theme color resolution と background fallback を解決する", () => {
    const computed = createComputedView(buildSource());
    const slide = getSlide(computed.slides, 0);

    // layout clrMapOvr が accent1 -> accent2 に差し替えるため、theme accent2 へ解決される。
    expect(slide.colorMap.accent1).toBe("accent2");
    expect(slide.background?.kind).toBe("fill");
    expect(slide.background?.sourceLayer).toBe("master");
    expect(slide.background?.kind === "fill" ? slide.background.fill.kind : undefined).toBe(
      "solid",
    );
    expect(
      slide.background?.kind === "fill" && slide.background.fill.kind === "solid"
        ? slide.background.fill.color
        : undefined,
    ).toEqual({ hex: "#336699", alpha: 1 });

    const slideTitle = findShape(slide.elements, "Slide title");
    expect(slideTitle.fill).toEqual(
      expect.objectContaining({
        kind: "solid",
        color: { hex: "#99b3cc", alpha: 0.5 },
      }),
    );

    const transformed = findShape(slide.elements, "Transformed colors");
    expect(transformed.fill).toEqual(
      expect.objectContaining({
        kind: "solid",
        color: { hex: "#ff8080", alpha: 1 },
      }),
    );
    expect(transformed.outline?.fill).toEqual(
      expect.objectContaining({
        kind: "solid",
        color: { hex: "#404040", alpha: 1 },
      }),
    );
  });

  it("placeholder matching と basic text style inheritance を解決する", () => {
    const computed = createComputedView(buildSource());
    const slideTitle = findShape(getSlide(computed.slides, 0).elements, "Slide title");

    expect(slideTitle.placeholderMatch?.layout?.name).toBe("Layout title");
    expect(slideTitle.placeholderMatch?.master?.name).toBe("Master title");
    expect(slideTitle.transform).toEqual({
      offsetX: 10,
      offsetY: 20,
      width: 300,
      height: 40,
    });
    expect(slideTitle.geometry).toEqual({ preset: "roundRect" });
    expect(slideTitle.textBody?.paragraphs[0].properties).toEqual({ align: "center" });
    expect(slideTitle.textBody?.paragraphs[0].runs[0]).toEqual({
      text: "Hello",
      properties: {
        bold: true,
        fontSize: 30,
        typeface: "Aptos",
        color: { hex: "#111111", alpha: 1 },
      },
    });
  });

  it("table の cell fill / text color / style default border を解決する", () => {
    const computed = createComputedView(buildSource());
    const table = findTable(getSlide(computed.slides, 0).elements, "Metrics table");
    const [cell] = table.table.rows[0].cells;

    expect(table.transform).toEqual({
      offsetX: 200,
      offsetY: 210,
      width: 220,
      height: 230,
    });
    expect(table.table.columns).toEqual([{ width: 500 }, { width: 600 }]);
    expect(cell).toMatchObject({
      gridSpan: 2,
      rowSpan: 1,
      hMerge: false,
      vMerge: false,
      fill: {
        kind: "solid",
        color: { hex: "#336699", alpha: 1 },
      },
    });
    expect(cell.textBody?.paragraphs[0].runs[0]).toEqual({
      text: "Q1",
      properties: {
        fontSize: 12,
        color: { hex: "#336699", alpha: 1 },
      },
    });
    expect(cell.borders?.top).toMatchObject({
      width: 12700,
      fill: {
        kind: "solid",
        color: { hex: "#000000", alpha: 1 },
      },
    });
  });

  it("inline table borders が table style default border より優先される", () => {
    const table = findTable(
      getSlide(createComputedView(buildSource()).slides, 0).elements,
      "Inline table",
    );
    const [cell] = table.table.rows[0].cells;

    expect(cell.borders?.left).toMatchObject({
      width: 999,
      fill: { kind: "solid", color: { hex: "#ff0000", alpha: 1 } },
    });
    expect(cell.borders?.top).toBeUndefined();
  });

  it("showMasterSp visibility と effective element ordering を解決する", () => {
    const computed = createComputedView(buildSource());

    expect(elementNames(getSlide(computed.slides, 0).elements)).toEqual([
      "Master decoration",
      "Layout decoration",
      "Slide title",
      "Hero image",
      "Transformed colors",
      "Metrics table",
      "Inline table",
    ]);

    // slide1 は showMasterSp=false なので master decoration が落ちる。
    expect(getSlide(computed.slides, 1).showMasterShapes).toBe(false);
    expect(getSlide(computed.slides, 1).layoutShowMasterShapes).toBe(true);
    expect(elementNames(getSlide(computed.slides, 1).elements)).toEqual([
      "Layout decoration",
      "Visible body",
    ]);

    const withoutVisibilityFilter = createComputedView(buildSource(), {
      applyMasterVisibility: false,
    });
    expect(elementNames(getSlide(withoutVisibilityFilter.slides, 1).elements)).toEqual([
      "Master decoration",
      "Layout decoration",
      "Visible body",
    ]);
  });

  it("connector / group / custom geometry を computed element として保持する", () => {
    const source = buildSource();
    const slidePath = asPartPath("ppt/slides/slide2.xml");
    const extended: CleanDocSource = {
      ...source,
      slides: source.slides.map((slide) =>
        slide.partPath === slidePath
          ? {
              ...slide,
              shapes: [
                ...slide.shapes,
                {
                  kind: "connector",
                  name: "Connector",
                  transform: transform(500, 510, 520, 530),
                  geometry: { preset: "bentConnector3", adjustValues: { adj1: 50000 } },
                  outline: {
                    width: asEmu(12700),
                    fill: { kind: "solid", color: { kind: "srgb", hex: "FF0000" } },
                    dashStyle: "dash",
                  },
                },
                {
                  kind: "group",
                  name: "Group",
                  transform: transform(600, 610, 620, 630),
                  childTransform: transform(0, 0, 620, 630),
                  children: [
                    shape("Custom child", {
                      transform: transform(1, 2, 3, 4),
                      geometry: {
                        kind: "custom",
                        paths: [{ width: 1000, height: 1000, commands: "M 0 0 L 1000 1000" }],
                      },
                    }),
                  ],
                },
              ],
            }
          : slide,
      ),
    };

    const slide = getSlide(createComputedView(extended).slides, 0);
    const connector = slide.elements.find(
      (element): element is ComputedConnectorElement => element.kind === "connector",
    );
    const group = slide.elements.find(
      (element): element is ComputedGroupElement => element.kind === "group",
    );

    expect(connector).toMatchObject({
      sourceNode: { name: "Connector" },
      transform: { offsetX: 500, offsetY: 510, width: 520, height: 530 },
      geometry: { preset: "bentConnector3", adjustValues: { adj1: 50000 } },
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { hex: "#ff0000", alpha: 1 } },
      },
    });
    expect(group).toMatchObject({
      sourceNode: { name: "Group" },
      transform: { offsetX: 600, offsetY: 610, width: 620, height: 630 },
      childTransform: { offsetX: 0, offsetY: 0, width: 620, height: 630 },
    });
    expect(group?.children[0]).toMatchObject({
      kind: "shape",
      geometry: {
        kind: "custom",
        paths: [{ width: 1000, height: 1000, commands: "M 0 0 L 1000 1000" }],
      },
    });
  });

  it("source model を in-place mutation しない", () => {
    const source = buildSource();
    const before = structuredClone(source);

    createComputedView(source);

    expect(source).toEqual(before);
  });

  it("target slide selection を presentation order の slide number で適用する", () => {
    const computed = createComputedView(buildSource(), { slides: [2] });

    expect(computed.slides.map((slide) => slide.partPath)).toEqual(["ppt/slides/slide1.xml"]);
    expect(getSlide(computed.slides, 0).slideNumber).toBe(2);
  });
});

function getSlide(slides: readonly ComputedSlide[], index: number): ComputedSlide {
  const slide = slides[index];
  if (slide === undefined) throw new Error(`slide at index ${index} not found`);
  return slide;
}

function findShape(elements: readonly ComputedElement[], name: string): ComputedShapeElement {
  const shape = elements.find(
    (element): element is ComputedShapeElement =>
      element.kind === "shape" && element.sourceNode.name === name,
  );
  if (shape === undefined) throw new Error(`shape '${name}' not found`);
  return shape;
}

function findTable(elements: readonly ComputedElement[], name: string): ComputedTableElement {
  const table = elements.find(
    (element): element is ComputedTableElement =>
      element.kind === "table" && element.sourceNode.name === name,
  );
  if (table === undefined) throw new Error(`table '${name}' not found`);
  return table;
}

function elementNames(elements: readonly ComputedElement[]): (string | undefined)[] {
  return elements.map((element) =>
    "name" in element.sourceNode ? element.sourceNode.name : undefined,
  );
}

function buildSource(): CleanDocSource {
  const masterTitle = placeholder("Master title", "title", 1, {
    transform: transform(100, 200, 300, 400),
    geometry: { preset: "rect" },
    textBody: textBody("", { fontSize: asPt(20), color: { kind: "srgb", hex: "222222" } }),
  });
  const layoutTitle = placeholder("Layout title", "ctrTitle", 1, {
    transform: transform(10, 20, 300, 40),
    geometry: { preset: "roundRect" },
    textBody: textBody("", {
      fontSize: asPt(30),
      typeface: "Aptos",
      color: { kind: "srgb", hex: "111111" },
    }),
  });

  return {
    packageGraph: {
      contentTypes: { defaults: [], overrides: [] },
      parts: [],
      relationships: [
        {
          sourcePartPath: asPartPath("ppt/slides/slide2.xml"),
          relationships: [
            {
              id: asRelationshipId("rIdImage"),
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              target: "../media/image1.png",
            },
            {
              id: asRelationshipId("rIdExternal"),
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
              target: "https://example.com/",
              targetMode: "External",
            },
          ],
        },
      ],
      media: [
        {
          partPath: asPartPath("ppt/media/image1.png"),
          contentType: "image/png",
          bytes: new Uint8Array([1, 2, 3]),
        },
      ],
    },
    presentation: {
      partPath: asPartPath("ppt/presentation.xml"),
      slideSize: { width: asEmu(9144000), height: asEmu(5143500) },
      slidePartPaths: [asPartPath("ppt/slides/slide2.xml"), asPartPath("ppt/slides/slide1.xml")],
    },
    slides: [
      {
        partPath: asPartPath("ppt/slides/slide1.xml"),
        layoutPartPath: asPartPath("ppt/slideLayouts/layout1.xml"),
        showMasterShapes: false,
        shapes: [
          placeholder("Empty placeholder", "body", 2, { textBody: textBody("") }),
          shape("Visible body", { transform: transform(1, 2, 3, 4) }),
        ],
      },
      {
        partPath: asPartPath("ppt/slides/slide2.xml"),
        layoutPartPath: asPartPath("ppt/slideLayouts/layout1.xml"),
        shapes: [
          placeholder("Slide title", "ctrTitle", 1, {
            textBody: textBody("Hello", { bold: true }),
            fill: {
              kind: "solid",
              color: {
                kind: "scheme",
                scheme: "accent1",
                transforms: [
                  { kind: "tint", value: asOoxmlPercent(50000) },
                  { kind: "alpha", value: asOoxmlPercent(50000) },
                ],
              },
            },
          }),
          {
            kind: "image",
            name: "Hero image",
            transform: transform(50, 60, 70, 80),
            blipRelationshipId: asRelationshipId("rIdImage"),
          },
          shape("Transformed colors", {
            transform: transform(11, 12, 13, 14),
            fill: {
              kind: "solid",
              color: {
                kind: "srgb",
                hex: "FF0000",
                transforms: [
                  { kind: "lumMod", value: asOoxmlPercent(50000) },
                  { kind: "lumOff", value: asOoxmlPercent(50000) },
                ],
              },
            },
            outline: {
              fill: {
                kind: "solid",
                color: {
                  kind: "srgb",
                  hex: "808080",
                  transforms: [{ kind: "shade", value: asOoxmlPercent(50000) }],
                },
              },
            },
          }),
          table("Metrics table", {
            transform: transform(200, 210, 220, 230),
            table: {
              tableStyleId: "{style}",
              columns: [{ width: asEmu(500) }, { width: asEmu(600) }],
              rows: [
                {
                  height: asEmu(700),
                  cells: [
                    {
                      textBody: textBody("Q1", {
                        fontSize: asPt(12),
                        color: { kind: "scheme", scheme: "accent1" },
                      }),
                      fill: { kind: "solid", color: { kind: "scheme", scheme: "accent1" } },
                      gridSpan: 2,
                      rowSpan: 1,
                      hMerge: false,
                      vMerge: false,
                    },
                  ],
                },
              ],
            },
          }),
          table("Inline table", {
            table: {
              tableStyleId: "{style}",
              columns: [{ width: asEmu(100) }],
              rows: [
                {
                  height: asEmu(100),
                  cells: [
                    {
                      borders: {
                        left: {
                          width: asEmu(999),
                          fill: { kind: "solid", color: { kind: "srgb", hex: "FF0000" } },
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
          }),
        ],
      },
    ],
    slideLayouts: [
      {
        partPath: asPartPath("ppt/slideLayouts/layout1.xml"),
        masterPartPath: asPartPath("ppt/slideMasters/master1.xml"),
        colorMapOverride: { mapping: { accent1: "accent2" } },
        shapes: [layoutTitle, shape("Layout decoration", { transform: transform(2, 2, 2, 2) })],
      },
    ],
    slideMasters: [
      {
        partPath: asPartPath("ppt/slideMasters/master1.xml"),
        themePartPath: asPartPath("ppt/theme/theme1.xml"),
        layoutPartPaths: [asPartPath("ppt/slideLayouts/layout1.xml")],
        colorMap: { mapping: { bg1: "lt1", tx1: "dk1", accent1: "accent1" } },
        background: {
          kind: "fill",
          fill: { kind: "solid", color: { kind: "scheme", scheme: "accent1" } },
        },
        shapes: [masterTitle, shape("Master decoration", { transform: transform(9, 9, 9, 9) })],
      },
    ],
    themes: [
      {
        partPath: asPartPath("ppt/theme/theme1.xml"),
        colorScheme: {
          colors: {
            lt1: { kind: "srgb", hex: "FFFFFF" },
            dk1: { kind: "srgb", hex: "000000" },
            accent1: { kind: "srgb", hex: "990000" },
            accent2: { kind: "srgb", hex: "336699" },
          },
        },
      },
    ],
    diagnostics: [],
  };
}

function buildSourceWithChartAndSmartArt(
  options: {
    readonly chartRelationshipType?: string;
  } = {},
): CleanDocSource {
  const source = buildSource();
  const slidePath = asPartPath("ppt/slides/slide2.xml");
  const dataPath = asPartPath("ppt/diagrams/data1.xml");
  const chartRelationshipType =
    options.chartRelationshipType ??
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      relationships: [
        ...source.packageGraph.relationships.map((entry) =>
          entry.sourcePartPath === slidePath
            ? {
                ...entry,
                relationships: [
                  ...entry.relationships,
                  {
                    id: asRelationshipId("rIdChart"),
                    type: chartRelationshipType,
                    target: "../charts/chart1.xml",
                  },
                  {
                    id: asRelationshipId("rIdDiagramData"),
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
                    target: "../diagrams/data1.xml",
                  },
                ],
              }
            : entry,
        ),
        {
          sourcePartPath: dataPath,
          relationships: [
            {
              id: asRelationshipId("rIdDiagramDrawing"),
              type: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
              target: "drawing1.xml",
            },
          ],
        },
      ],
      rawParts: [
        rawXmlPart(
          "ppt/charts/chart1.xml",
          `<c:chartSpace><c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>`,
        ),
        rawXmlPart("ppt/diagrams/data1.xml", `<dgm:dataModel/>`),
        rawXmlPart("ppt/diagrams/drawing1.xml", `<dsp:drawing><dsp:spTree/></dsp:drawing>`),
      ],
    },
    slides: source.slides.map((slide) =>
      slide.partPath === slidePath
        ? {
            ...slide,
            shapes: [
              ...slide.shapes,
              {
                kind: "chart",
                name: "Revenue chart",
                transform: transform(300, 310, 320, 330),
                chartRelationshipId: asRelationshipId("rIdChart"),
              },
              {
                kind: "smartArt",
                name: "Process",
                transform: transform(400, 410, 420, 430),
                dataRelationshipId: asRelationshipId("rIdDiagramData"),
              },
            ],
          }
        : slide,
    ),
  };
}

function rawXmlPart(partPath: string, xml: string) {
  return {
    kind: "binary" as const,
    partPath: asPartPath(partPath),
    contentType: "application/xml",
    bytes: new TextEncoder().encode(xml),
  };
}

function transform(offsetX: number, offsetY: number, width: number, height: number) {
  return {
    offsetX: asEmu(offsetX),
    offsetY: asEmu(offsetY),
    width: asEmu(width),
    height: asEmu(height),
  };
}

function shape(name: string, overrides: Partial<SourceShape> = {}): SourceShape {
  return {
    kind: "shape",
    name,
    ...overrides,
  };
}

function table(name: string, overrides: Omit<SourceTable, "kind" | "name">): SourceTable {
  return {
    kind: "table",
    name,
    ...overrides,
  };
}

function placeholder(
  name: string,
  type: string,
  index: number,
  overrides: Partial<SourceShape> = {},
): SourceShape {
  return shape(name, {
    placeholder: { type, index },
    ...overrides,
  });
}

function textBody(text: string, runProperties: SourceRunProperties = {}): SourceTextBody {
  return {
    paragraphs: [
      {
        properties: { align: "center" as const },
        runs: [{ kind: "textRun" as const, text, properties: runProperties }],
      },
    ],
  };
}
