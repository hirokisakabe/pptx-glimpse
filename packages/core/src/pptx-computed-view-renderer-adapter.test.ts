import type {
  PptxComputedView,
  PptxSourceModel,
  SourceChart,
  SourceShape,
  SourceTable,
} from "@pptx-glimpse/document";
import {
  asEmu,
  asOoxmlAngle,
  asOoxmlPercent,
  asPartPath,
  asPt,
  asRawSidecarId,
  asRelationshipId,
  createComputedView,
} from "@pptx-glimpse/document";
import type { SlideElement } from "@pptx-glimpse/renderer";
import { describe, expect, it } from "vitest";

import { adaptComputedViewToRendererModel } from "./pptx-computed-view-renderer-adapter.js";
import { unsafeFixtureAssertion } from "./unsafe-type-assertion.js";

describe("adaptComputedViewToRendererModel", () => {
  it("covers pptx-computed-view-renderer-adapter behavior 1", () => {
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

  it("covers pptx-computed-view-renderer-adapter behavior 2", () => {
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
        wrap: "none",
        autoFit: "normAutofit",
        fontScale: 0.625,
        lnSpcReduction: 0.2,
        numCol: 2,
        vert: "eaVert",
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
                fontFamilyEa: "Yu Gothic",
                fontFamilyCs: "Arial",
                bold: true,
                italic: false,
                underline: false,
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

  it("covers pptx-computed-view-renderer-adapter behavior 3", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(
        buildSource({
          extraSlideShapes: [
            shape("Gradient fill", {
              transform: transform(30, 31, 32, 33),
              fill: {
                kind: "gradient",
                gradientType: "linear",
                angle: asOoxmlAngle(5400000),
                stops: [
                  { position: 0, color: { kind: "scheme", scheme: "accent1" } },
                  { position: 1, color: { kind: "srgb", hex: "00FF00" } },
                ],
              },
            }),
            shape("Pattern fill", {
              transform: transform(40, 41, 42, 43),
              fill: {
                kind: "pattern",
                preset: "pct20",
                foregroundColor: { kind: "scheme", scheme: "accent1" },
                backgroundColor: { kind: "srgb", hex: "FFFFFF" },
              },
            }),
            shape("Image fill", {
              transform: transform(50, 51, 52, 53),
              fill: {
                kind: "image",
                blipRelationshipId: asRelationshipId("rIdImage"),
                tile: {
                  tx: asEmu(1),
                  ty: asEmu(2),
                  sx: 0.5,
                  sy: 0.75,
                  flip: "xy",
                  align: "ctr",
                },
              },
            }),
          ],
        }),
      ),
    );

    expect(findElementByAltText(result.slides[0].elements, "Gradient fill")).toMatchObject({
      fill: {
        type: "gradient",
        angle: 90,
        stops: [
          { position: 0, color: { hex: "#336699", alpha: 1 } },
          { position: 1, color: { hex: "#00ff00", alpha: 1 } },
        ],
      },
    });
    expect(findElementByAltText(result.slides[0].elements, "Pattern fill")).toMatchObject({
      fill: {
        type: "pattern",
        preset: "pct20",
        foregroundColor: { hex: "#336699", alpha: 1 },
        backgroundColor: { hex: "#ffffff", alpha: 1 },
      },
    });
    expect(findElementByAltText(result.slides[0].elements, "Image fill")).toMatchObject({
      fill: {
        type: "image",
        imageData: "AQID",
        mimeType: "image/png",
        tile: {
          tx: 1,
          ty: 2,
          sx: 0.5,
          sy: 0.75,
          flip: "xy",
          align: "ctr",
        },
      },
    });
    expect(result.diagnostics).toEqual([]);
  });

  it("covers pptx-computed-view-renderer-adapter behavior 4", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(
        buildSource({
          extraSlideShapes: [
            shape("Shadowed shape", {
              transform: transform(90, 91, 92, 93),
              effects: {
                outerShadow: {
                  blurRadius: asEmu(10),
                  distance: asEmu(20),
                  direction: asOoxmlAngle(5400000),
                  color: { kind: "scheme", scheme: "accent1" },
                  alignment: "ctr",
                  rotateWithShape: false,
                },
                softEdge: { radius: asEmu(30) },
              },
            }),
            {
              kind: "image",
              name: "Effect image",
              transform: transform(100, 101, 102, 103),
              blipRelationshipId: asRelationshipId("rIdImage"),
              effects: { glow: { radius: asEmu(40), color: { kind: "srgb", hex: "FFFFFF" } } },
              blipEffects: {
                grayscale: true,
                biLevel: { threshold: 0.25 },
                blur: { radius: asEmu(50), grow: false },
                lum: { brightness: 0.1, contrast: -0.2 },
                duotone: {
                  color1: { kind: "srgb", hex: "000000" },
                  color2: { kind: "scheme", scheme: "accent1" },
                },
                clrChange: {
                  from: { kind: "srgb", hex: "FF0000" },
                  to: { kind: "scheme", scheme: "accent2" },
                },
              },
            },
          ],
        }),
      ),
    );

    expect(findElementByAltText(result.slides[0].elements, "Shadowed shape")).toMatchObject({
      effects: {
        outerShadow: {
          blurRadius: 10,
          distance: 20,
          direction: 90,
          color: { hex: "#336699", alpha: 1 },
          alignment: "ctr",
          rotateWithShape: false,
        },
        softEdge: { radius: 30 },
      },
    });
    expect(findElementByAltText(result.slides[0].elements, "Effect image")).toMatchObject({
      effects: {
        glow: { radius: 40, color: { hex: "#ffffff", alpha: 1 } },
      },
      blipEffects: {
        grayscale: true,
        biLevel: { threshold: 0.25 },
        blur: { radius: 50, grow: false },
        lum: { brightness: 0.1, contrast: -0.2 },
        duotone: {
          color1: { hex: "#000000", alpha: 1 },
          color2: { hex: "#336699", alpha: 1 },
        },
        clrChange: {
          clrFrom: { hex: "#ff0000", alpha: 1 },
          clrTo: { hex: "#336699", alpha: 1 },
        },
      },
    });
    expect(result.diagnostics).toEqual([]);
  });

  it("covers pptx-computed-view-renderer-adapter behavior 5", () => {
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

  it("covers pptx-computed-view-renderer-adapter behavior 6", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(
        buildSource({
          extraSlideShapes: [
            {
              kind: "connector",
              name: "Connector",
              transform: transform(500, 510, 520, 530),
              geometry: { preset: "bentConnector3", adjustValues: { adj1: 50000 } },
              outline: {
                width: asEmu(25400),
                fill: { kind: "solid", color: { kind: "srgb", hex: "FF0000" } },
                dashStyle: "dash",
                tailEnd: { type: "triangle", width: "med", length: "med" },
              },
              effects: {
                outerShadow: {
                  blurRadius: asEmu(10),
                  distance: asEmu(20),
                  direction: asOoxmlAngle(5400000),
                  color: { kind: "srgb", hex: "000000" },
                  alignment: "b",
                  rotateWithShape: true,
                },
              },
            },
            {
              kind: "group",
              name: "Group",
              transform: transform(600, 610, 620, 630),
              childTransform: transform(0, 0, 620, 630),
              effects: { softEdge: { radius: asEmu(30) } },
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
        }),
      ),
    );

    const connector = findElementByAltText(result.slides[0].elements, "Connector");
    const group = findElementByAltText(result.slides[0].elements, "Group");
    expect(connector).toMatchObject({
      type: "connector",
      transform: { offsetX: 500, offsetY: 510, extentWidth: 520, extentHeight: 530 },
      geometry: { type: "preset", preset: "bentConnector3", adjustValues: { adj1: 50000 } },
      outline: {
        width: 25400,
        fill: { type: "solid", color: { hex: "#ff0000", alpha: 1 } },
        dashStyle: "dash",
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
      effects: {
        outerShadow: {
          blurRadius: 10,
          distance: 20,
          direction: 90,
          color: { hex: "#000000", alpha: 1 },
        },
      },
    });
    expect(group).toMatchObject({
      type: "group",
      transform: { offsetX: 600, offsetY: 610, extentWidth: 620, extentHeight: 630 },
      childTransform: { offsetX: 0, offsetY: 0, extentWidth: 620, extentHeight: 630 },
      effects: { softEdge: { radius: 30 } },
      children: [
        {
          type: "shape",
          geometry: {
            type: "custom",
            paths: [{ width: 1000, height: 1000, commands: "M 0 0 L 1000 1000" }],
          },
        },
      ],
    });
    expect(result.diagnostics).toEqual([]);
  });

  it("covers pptx-computed-view-renderer-adapter behavior 7", () => {
    const source = buildSource({
      extraSlideShapes: [
        shape("EMF fill", {
          transform: transform(110, 111, 112, 113),
          fill: { kind: "image", blipRelationshipId: asRelationshipId("rIdImage") },
        }),
      ],
    });
    const sourceWithEmfMedia: PptxSourceModel = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        media: source.packageGraph.media.map((media) => ({
          ...media,
          contentType: "image/x-emf",
        })),
      },
    };
    const result = adaptComputedViewToRendererModel(createComputedView(sourceWithEmfMedia));

    expect(findElementByAltText(result.slides[0].elements, "Hero image")).toMatchObject({
      type: "image",
      mimeType: "image/emf",
    });
    expect(findElementByAltText(result.slides[0].elements, "EMF fill")).toMatchObject({
      type: "shape",
      fill: { type: "image", mimeType: "image/emf" },
    });
  });

  it("covers pptx-computed-view-renderer-adapter behavior 8", () => {
    const source = buildSource();
    const sourceWithUnknownMedia: PptxSourceModel = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        media: source.packageGraph.media.map((media) => ({
          ...media,
          contentType: "application/octet-stream",
        })),
      },
    };

    const result = adaptComputedViewToRendererModel(createComputedView(sourceWithUnknownMedia));

    expect(findElementByAltText(result.slides[0].elements, "Hero image")).toMatchObject({
      type: "image",
      mimeType: "image/png",
    });
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: "pptx-computed-view-adapter.unsupported-image-mime-type",
        sourcePartPath: "ppt/slides/slide1.xml",
      }),
    );
  });

  it("covers pptx-computed-view-renderer-adapter behavior 9", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(
        buildSource({
          extraSlideShapes: [
            shape("Unknown shadow alignment", {
              transform: transform(110, 111, 112, 113),
              effects: {
                outerShadow: {
                  blurRadius: asEmu(10),
                  distance: asEmu(20),
                  direction: asOoxmlAngle(30),
                  color: { kind: "srgb", hex: "000000" },
                  alignment: unsafeFixtureAssertion<"b">("unknown"),
                  rotateWithShape: true,
                },
              },
            }),
            {
              kind: "image",
              name: "Unknown tile alignment",
              transform: transform(120, 121, 122, 123),
              blipRelationshipId: asRelationshipId("rIdImage"),
              tile: {
                tx: asEmu(1),
                ty: asEmu(2),
                sx: 0.5,
                sy: 0.75,
                flip: "none",
                align: unsafeFixtureAssertion<"tl">("unknown"),
              },
            },
          ],
        }),
      ),
    );

    expect(
      findElementByAltText(result.slides[0].elements, "Unknown shadow alignment"),
    ).toMatchObject({
      type: "shape",
      effects: { outerShadow: { alignment: "b" } },
    });
    expect(findElementByAltText(result.slides[0].elements, "Unknown tile alignment")).toMatchObject(
      {
        type: "image",
        tile: { align: "tl" },
      },
    );
    expect(result.diagnostics.map((diagnostic) => diagnostic.code)).toEqual(
      expect.arrayContaining(["pptx-computed-view-adapter.unsupported-rectangle-alignment"]),
    );
  });

  it("covers pptx-computed-view-renderer-adapter behavior 10", () => {
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
        offsetX: 100,
        offsetY: 110,
        extentWidth: 420,
        extentHeight: 430,
      },
    });
    expect(smartArt?.children.map((element) => element.type)).toEqual(["image", "group", "shape"]);
    expect(smartArt?.children[0]).toMatchObject({
      type: "image",
      transform: {
        offsetX: 11,
        offsetY: 12,
        extentWidth: 13,
        extentHeight: 14,
      },
      mimeType: "image/png",
    });
    expect(smartArt?.children[1]).toMatchObject({
      type: "group",
      childTransform: {
        offsetX: 210,
        offsetY: 220,
        extentWidth: 230,
        extentHeight: 240,
      },
      children: [
        expect.objectContaining({
          type: "shape",
          fill: { type: "solid", color: { hex: "#00ff00", alpha: 1 } },
        }),
      ],
    });
    expect(smartArt?.children[2]).toMatchObject({
      type: "shape",
      fill: { type: "solid", color: { hex: "#336699", alpha: 1 } },
      outline: {
        fill: { type: "solid", color: { hex: "#123456", alpha: 1 } },
        width: 12700,
      },
      textBody: {
        paragraphs: [
          expect.objectContaining({
            runs: [expect.objectContaining({ text: "Step" })],
          }),
        ],
      },
    });
    expect(result.diagnostics).toEqual([]);
  });

  it("covers pptx-computed-view-renderer-adapter behavior 11", () => {
    const result = adaptComputedViewToRendererModel(buildComputedViewWithChartData());
    const chart = result.slides[0].elements.find((element) => element.type === "chart");

    expect(chart).toMatchObject({
      type: "chart",
      transform: {
        offsetX: 10,
        offsetY: 20,
        extentWidth: 300,
        extentHeight: 200,
      },
      chart: {
        chartType: "bubble",
        title: "Pipeline",
        categories: ["A", "B"],
        series: [
          {
            name: "Weighted",
            values: [4, 8],
            xValues: [1, 2],
            bubbleSizes: [10, 20],
            color: { hex: "#336699", alpha: 0.75 },
          },
        ],
        legend: { position: "tr" },
      },
    });
    expect(result.diagnostics).toEqual([]);
  });

  it("covers pptx-computed-view-renderer-adapter behavior 12", () => {
    const result = adaptComputedViewToRendererModel(
      createComputedView(
        buildSourceWithChartAndSmartArt({
          smartArtDrawingXml: `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"/>`,
        }),
      ),
    );

    expect(
      result.slides[0].elements.some(
        (element) => element.type === "group" && element.altText === "Process",
      ),
    ).toBe(false);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: "pptx-computed-view-adapter.unresolved-smartart-skipped",
        sourcePartPath: "ppt/diagrams/drawing1.xml",
      }),
    );
  });

  it("covers pptx-computed-view-renderer-adapter behavior 13", () => {
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
        code: "pptx-computed-view-adapter.raw-element-skipped",
        slideNumber: 1,
        sourcePartPath: "ppt/slides/slide1.xml",
      }),
    );
  });

  it("covers pptx-computed-view-renderer-adapter behavior 14", () => {
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
        "pptx-computed-view-adapter.raw-background-ignored",
        "pptx-computed-view-adapter.missing-transform",
        "pptx-computed-view-adapter.raw-fill-ignored",
        "pptx-computed-view-adapter.unresolved-image-skipped",
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
  readonly extraSlideShapes?: PptxSourceModel["slides"][number]["shapes"];
  readonly slideBackground?: PptxSourceModel["slides"][number]["background"];
}

function buildSource(options: BuildSourceOptions = {}): PptxSourceModel {
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
              properties: {
                marginLeft: asEmu(111),
                anchor: "middle",
                wrap: "none",
                autoFit: "normAutofit",
                fontScale: 0.625,
                lnSpcReduction: 0.2,
                numCol: 2,
                vert: "eaVert",
              },
              paragraphs: [
                {
                  properties: {
                    align: "center",
                    lineSpacing: { type: "pts", value: 1200 },
                    level: 2,
                  },
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
            textBody: listStyleTextBody({
              fontSize: asPt(30),
              typeface: "Aptos",
              typefaceEa: "+mn-ea",
              typefaceCs: "+mn-cs",
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
        fontScheme: {
          minorEastAsian: "Yu Gothic",
          minorComplexScript: "Arial",
        },
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
  options: { readonly smartArtDrawingXml?: string } = {},
): PptxSourceModel {
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
        {
          sourcePartPath: asPartPath("ppt/diagrams/drawing1.xml"),
          relationships: [
            {
              id: asRelationshipId("rIdDiagramImage"),
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              target: "../media/image1.png",
            },
          ],
        },
      ],
      rawParts: [
        rawXmlPart("ppt/charts/chart1.xml", chartXml()),
        rawXmlPart("ppt/diagrams/data1.xml", `<dgm:dataModel/>`),
        rawXmlPart("ppt/diagrams/drawing1.xml", options.smartArtDrawingXml ?? smartArtDrawingXml()),
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

function buildComputedViewWithChartData(): PptxComputedView {
  const sourceNode: SourceChart = {
    kind: "chart",
    name: "Pipeline chart",
  };
  return {
    slides: [
      {
        slideNumber: 1,
        partPath: asPartPath("ppt/slides/slide1.xml"),
        relationships: [],
        colorMap: {},
        colorScheme: {},
        showMasterShapes: true,
        layoutShowMasterShapes: true,
        elements: [
          {
            kind: "chart",
            sourceLayer: "slide",
            sourcePartPath: asPartPath("ppt/slides/slide1.xml"),
            sourceNode,
            transform: transform(10, 20, 300, 200),
            chartData: {
              chartType: "bubble",
              title: "Pipeline",
              categories: ["A", "B"],
              series: [
                {
                  name: "Weighted",
                  values: [4, 8],
                  xValues: [1, 2],
                  bubbleSizes: [10, 20],
                  color: { hex: "#336699", alpha: 0.75 },
                },
              ],
              legend: { position: "tr" },
            },
          },
        ],
      },
    ],
  };
}

function smartArtDrawingXml(): string {
  return `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dsp:spTree>
    <dsp:grpSpPr><a:xfrm><a:off x="700" y="710"/><a:ext cx="720" cy="730"/><a:chOff x="100" y="110"/><a:chExt cx="420" cy="430"/></a:xfrm></dsp:grpSpPr>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="3" name="Diagram image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdDiagramImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="11" y="12"/><a:ext cx="13" cy="14"/></a:xfrm></p:spPr>
    </p:pic>
    <p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="4" name="Diagram group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="200" y="201"/><a:ext cx="202" cy="203"/><a:chOff x="210" y="220"/><a:chExt cx="230" cy="240"/></a:xfrm><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="Grouped shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="21" y="22"/><a:ext cx="23" cy="24"/></a:xfrm><a:grpFill/></p:spPr>
      </p:sp>
    </p:grpSp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="1" name="Step"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></p:spPr>
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

function listStyleTextBody(
  runProperties: NonNullable<
    SourceShape["textBody"]
  >["paragraphs"][number]["runs"][number]["properties"],
): NonNullable<SourceShape["textBody"]> {
  return {
    listStyle: {
      defaultParagraph: { defaultRunProperties: runProperties },
      levels: [],
    },
    paragraphs: [{ runs: [] }],
  };
}

function rawSidecar(id: string, name: string) {
  return {
    id: asRawSidecarId(id),
    node: { name },
  };
}
