import type { SlideElement } from "pptx-glimpse-renderer";
import { describe, expect, it } from "vitest";

import type {
  CleanDocSource,
  SourceShape,
  SourceTable,
} from "../../pptx-glimpse-document/src/experimental.js";
import {
  asEmu,
  asOoxmlAngle,
  asOoxmlPercent,
  asPartPath,
  asPt,
  asRawSidecarId,
  asRelationshipId,
  createComputedView,
} from "../../pptx-glimpse-document/src/experimental.js";
import { adaptComputedViewToRendererModel } from "./cleandoc-renderer-adapter.js";

describe("adaptComputedViewToRendererModel", () => {
  it("slide size / background / effective element ordering を renderer model に変換する", () => {
    const result = adaptComputedViewToRendererModel(createComputedView(buildSource()));

    expect(result.slideSize).toEqual({ width: 9144000, height: 5143500 });
    expect(result.slides).toHaveLength(1);
    const slide = result.slides[0];

    expect(slide.slideNumber).toBe(1);
    expect(slide.background?.fill).toEqual({
      type: "solid",
      color: { hex: "#336699", alpha: 1 },
    });
    expect(slide.elements.map(getAltText)).toEqual([
      "Master decoration",
      "Layout decoration",
      "Slide title",
      "Hero image",
    ]);
    expect(result.diagnostics).toEqual([]);
  });

  it("simple shape / text / image と theme-resolved fill / outline を変換する", () => {
    const result = adaptComputedViewToRendererModel(createComputedView(buildSource()));
    const slide = result.slides[0];
    const title = findElementByAltText(slide.elements, "Slide title");
    const image = findElementByAltText(slide.elements, "Hero image");

    expect(title).toMatchObject({
      type: "shape",
      transform: {
        offsetX: 10,
        offsetY: 20,
        extentWidth: 300,
        extentHeight: 40,
        rotation: 2,
        flipH: true,
        flipV: false,
      },
      geometry: { type: "preset", preset: "roundRect", adjustValues: {} },
      fill: { type: "solid", color: { hex: "#99b3cc", alpha: 0.5 } },
      outline: {
        width: 12700,
        fill: { type: "solid", color: { hex: "#2f528f", alpha: 1 } },
      },
      placeholderType: "ctrTitle",
      placeholderIdx: 1,
    });
    expect(title.type === "shape" ? title.textBody : undefined).toMatchObject({
      bodyProperties: {
        anchor: "ctr",
        marginLeft: 111,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
      },
      paragraphs: [
        {
          properties: {
            alignment: "ctr",
            lineSpacing: { type: "pts", value: 1200 },
            level: 2,
          },
          runs: [
            {
              text: "Hello",
              properties: {
                fontSize: 30,
                fontFamily: "Aptos",
                bold: true,
                italic: false,
                underline: true,
                color: { hex: "#111111", alpha: 1 },
              },
            },
          ],
        },
      ],
    });

    expect(image).toMatchObject({
      type: "image",
      transform: {
        offsetX: 50,
        offsetY: 60,
        extentWidth: 70,
        extentHeight: 80,
      },
      imageData: "AQID",
      mimeType: "image/png",
      srcRect: { left: 0.1, top: 0, right: 0.2, bottom: 0 },
    });
  });

  it("table を renderer model に変換する", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(
        buildSource({
          extraSlideShapes: [
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
              transform: transform(300, 310, 320, 330),
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
        }),
      ),
    );

    const [metrics, inline] = result.slides[0].elements.filter(
      (element): element is Extract<SlideElement, { type: "table" }> => element.type === "table",
    );
    expect(metrics).toMatchObject({
      type: "table",
      transform: {
        offsetX: 200,
        offsetY: 210,
        extentWidth: 220,
        extentHeight: 230,
      },
      table: {
        columns: [{ width: 500 }, { width: 600 }],
        rows: [
          {
            height: 700,
            cells: [
              {
                textBody: {
                  paragraphs: [
                    {
                      runs: [
                        {
                          text: "Q1",
                          properties: {
                            fontSize: 12,
                            color: { hex: "#336699", alpha: 1 },
                          },
                        },
                      ],
                    },
                  ],
                },
                fill: { type: "solid", color: { hex: "#336699", alpha: 1 } },
                borders: {
                  top: {
                    width: 12700,
                    fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
                  },
                },
                gridSpan: 2,
              },
            ],
          },
        ],
      },
    });

    expect(inline?.table.rows[0].cells[0].borders).toEqual({
      top: null,
      bottom: null,
      left: {
        width: 999,
        fill: { type: "solid", color: { hex: "#ff0000", alpha: 1 } },
        dashStyle: "solid",
        headEnd: null,
        tailEnd: null,
      },
      right: null,
    });
    expect(result.diagnostics).toEqual([]);
  });

  it("chart と SmartArt fallback を renderer model に変換する", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(buildSourceWithChartAndSmartArt()),
    );
    const chart = result.slides[0].elements.find((element) => element.type === "chart");
    const smartArt = result.slides[0].elements.find(
      (element): element is Extract<SlideElement, { type: "group" }> =>
        element.type === "group" && element.altText === "Process",
    );

    expect(chart).toMatchObject({
      type: "chart",
      transform: {
        offsetX: 300,
        offsetY: 310,
        extentWidth: 320,
        extentHeight: 330,
      },
      chart: {
        chartType: "bar",
        categories: ["Q1", "Q2"],
        series: [{ name: "Sales", values: [4, 7] }],
      },
    });
    expect(smartArt).toMatchObject({
      type: "group",
      transform: {
        offsetX: 400,
        offsetY: 410,
        extentWidth: 420,
        extentHeight: 430,
      },
      childTransform: {
        offsetX: 0,
        offsetY: 0,
        extentWidth: 420,
        extentHeight: 430,
      },
    });
    expect(smartArt?.children.map((element) => element.type)).toEqual(["shape"]);
    expect(result.diagnostics).toEqual([]);
  });

  it("unsupported raw elements を diagnostic として扱い renderer model に漏らさない", () => {
    const source = buildSource({
      extraSlideShapes: [
        {
          kind: "raw",
          raw: rawSidecar("raw-1", "p:graphicFrame"),
        },
      ],
    });

    const result = adaptComputedViewToRendererModel(createComputedView(source));

    expect(result.slides[0]?.elements).toHaveLength(4);
    expect(result.slides[0]?.elements.map(getAltText)).toEqual([
      "Master decoration",
      "Layout decoration",
      "Slide title",
      "Hero image",
    ]);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        severity: "warning",
        code: "cleandoc-adapter.raw-element-skipped",
        slideNumber: 1,
        sourcePartPath: "ppt/slides/slide1.xml",
      }),
    );
  });

  it("fallback が必要な unsupported subset を diagnostic に残す", () => {
    const source = buildSource({
      slideBackground: { kind: "raw", raw: rawSidecar("raw-bg", "p:bg") },
      extraSlideShapes: [
        shape("Missing transform"),
        shape("Raw fill shape", {
          transform: transform(90, 91, 92, 93),
          fill: { kind: "raw", raw: rawSidecar("raw-fill", "a:gradFill") },
        }),
        {
          kind: "image",
          name: "Missing media",
          transform: transform(100, 101, 102, 103),
          blipRelationshipId: asRelationshipId("rIdMissing"),
        },
      ],
    });

    const result = adaptComputedViewToRendererModel(createComputedView(source));

    expect(result.slides[0]?.background).toBeNull();
    expect(findElementByAltText(result.slides[0].elements, "Missing transform")).toMatchObject({
      transform: {
        offsetX: 0,
        offsetY: 0,
        extentWidth: 0,
        extentHeight: 0,
      },
    });
    expect(findElementByAltText(result.slides[0].elements, "Raw fill shape")).toMatchObject({
      fill: null,
    });
    expect(result.slides[0]?.elements.map(getAltText)).not.toContain("Missing media");
    expect(result.diagnostics.map((diagnostic) => diagnostic.code)).toEqual(
      expect.arrayContaining([
        "cleandoc-adapter.raw-background-ignored",
        "cleandoc-adapter.missing-transform",
        "cleandoc-adapter.raw-fill-ignored",
        "cleandoc-adapter.unresolved-image-skipped",
      ]),
    );
  });
});

function findElementByAltText(elements: readonly SlideElement[], altText: string): SlideElement {
  const element = elements.find((candidate) => getAltText(candidate) === altText);
  if (element === undefined) throw new Error(`element '${altText}' not found`);
  return element;
}

function getAltText(element: SlideElement): string | undefined {
  return "altText" in element ? element.altText : undefined;
}

interface BuildSourceOptions {
  readonly extraSlideShapes?: CleanDocSource["slides"][number]["shapes"];
  readonly slideBackground?: CleanDocSource["slides"][number]["background"];
}

function buildSource(options: BuildSourceOptions = {}): CleanDocSource {
  const slidePath = asPartPath("ppt/slides/slide1.xml");
  const layoutPath = asPartPath("ppt/slideLayouts/layout1.xml");
  const masterPath = asPartPath("ppt/slideMasters/master1.xml");
  const themePath = asPartPath("ppt/theme/theme1.xml");

  return {
    packageGraph: {
      contentTypes: { defaults: [], overrides: [] },
      parts: [],
      relationships: [
        {
          sourcePartPath: slidePath,
          relationships: [
            {
              id: asRelationshipId("rIdImage"),
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              target: "../media/image1.png",
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
      slidePartPaths: [slidePath],
    },
    slides: [
      {
        partPath: slidePath,
        layoutPartPath: layoutPath,
        ...(options.slideBackground !== undefined ? { background: options.slideBackground } : {}),
        shapes: [
          {
            kind: "shape",
            name: "Slide title",
            placeholder: { type: "ctrTitle", index: 1 },
            textBody: {
              properties: { marginLeft: asEmu(111), anchor: "middle" },
              paragraphs: [
                {
                  properties: { align: "center", lineSpacingPts: 1200, level: 2 },
                  runs: [{ kind: "textRun", text: "Hello", properties: { bold: true } }],
                },
              ],
            },
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
            outline: {
              width: asEmu(12700),
              fill: { kind: "solid", color: { kind: "srgb", hex: "2F528F" } },
            },
          },
          {
            kind: "image",
            name: "Hero image",
            transform: transform(50, 60, 70, 80),
            blipRelationshipId: asRelationshipId("rIdImage"),
            crop: {
              left: asOoxmlPercent(10000),
              right: asOoxmlPercent(20000),
            },
          },
          ...(options.extraSlideShapes ?? []),
        ],
      },
    ],
    slideLayouts: [
      {
        partPath: layoutPath,
        masterPartPath: masterPath,
        colorMapOverride: { mapping: { accent1: "accent2" } },
        shapes: [
          placeholder("Layout title", "ctrTitle", 1, {
            transform: transform(10, 20, 300, 40, { rotation: 120000, flipHorizontal: true }),
            geometry: { preset: "roundRect" },
            textBody: textBody("", {
              fontSize: asPt(30),
              typeface: "Aptos",
              color: { kind: "srgb", hex: "111111" },
              underline: true,
            }),
          }),
          shape("Layout decoration", { transform: transform(2, 2, 2, 2) }),
        ],
      },
    ],
    slideMasters: [
      {
        partPath: masterPath,
        themePartPath: themePath,
        layoutPartPaths: [layoutPath],
        colorMap: { mapping: { bg1: "lt1", tx1: "dk1", accent1: "accent1" } },
        background: {
          kind: "fill",
          fill: { kind: "solid", color: { kind: "scheme", scheme: "accent1" } },
        },
        shapes: [shape("Master decoration", { transform: transform(1, 1, 1, 1) })],
      },
    ],
    themes: [
      {
        partPath: themePath,
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

function buildSourceWithChartAndSmartArt(): CleanDocSource {
  const source = buildSource({
    extraSlideShapes: [
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
  });
  const slidePath = asPartPath("ppt/slides/slide1.xml");
  const dataPath = asPartPath("ppt/diagrams/data1.xml");

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
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
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
        rawXmlPart("ppt/charts/chart1.xml", chartXml()),
        rawXmlPart("ppt/diagrams/data1.xml", `<dgm:dataModel/>`),
        rawXmlPart("ppt/diagrams/drawing1.xml", smartArtDrawingXml()),
      ],
    },
  };
}

function chartXml(): string {
  return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart>
    <c:ser>
      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:barChart></c:plotArea></c:chart>
</c:chartSpace>`;
}

function smartArtDrawingXml(): string {
  return `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <dsp:spTree>
    <dsp:grpSpPr><a:xfrm><a:chOff x="0" y="0"/><a:chExt cx="420" cy="430"/></a:xfrm></dsp:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="1" name="Step"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>Step</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </dsp:spTree>
</dsp:drawing>`;
}

function rawXmlPart(partPath: string, xml: string) {
  return {
    kind: "binary" as const,
    partPath: asPartPath(partPath),
    contentType: "application/xml",
    bytes: new TextEncoder().encode(xml),
  };
}

function transform(
  offsetX: number,
  offsetY: number,
  width: number,
  height: number,
  options: { readonly rotation?: number; readonly flipHorizontal?: boolean } = {},
) {
  return {
    offsetX: asEmu(offsetX),
    offsetY: asEmu(offsetY),
    width: asEmu(width),
    height: asEmu(height),
    ...(options.rotation !== undefined ? { rotation: asOoxmlAngle(options.rotation) } : {}),
    ...(options.flipHorizontal !== undefined ? { flipHorizontal: options.flipHorizontal } : {}),
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

function textBody(
  text: string,
  runProperties: NonNullable<
    SourceShape["textBody"]
  >["paragraphs"][number]["runs"][number]["properties"],
): NonNullable<SourceShape["textBody"]> {
  return {
    paragraphs: [
      {
        runs: [{ kind: "textRun", text, properties: runProperties }],
      },
    ],
  };
}

function rawSidecar(id: string, name: string) {
  return {
    id: asRawSidecarId(id),
    node: { name },
  };
}
