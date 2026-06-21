import type { CleanDocSource, SourceShape } from "@pptx-glimpse/document/experimental";
import {
  asEmu,
  asOoxmlAngle,
  asOoxmlPercent,
  asPartPath,
  asPt,
  asRelationshipId,
  createComputedView,
} from "@pptx-glimpse/document/experimental";
import type { SlideElement } from "pptx-glimpse-renderer";
import { describe, expect, it } from "vitest";

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
    expect(title?.type === "shape" ? title.textBody : undefined).toMatchObject({
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

  it("unsupported raw elements are diagnosed instead of leaking into renderer model", () => {
    const source = buildSource({
      extraSlideShapes: [
        {
          kind: "raw",
          raw: { id: "raw-1", xml: { name: "p:graphicFrame" } },
        },
      ],
    });

    const result = adaptComputedViewToRendererModel(createComputedView(source));

    expect(result.slides[0]?.elements.map(getAltText)).not.toContain(undefined);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        severity: "warning",
        code: "cleandoc-adapter.raw-element-skipped",
        slideNumber: 1,
        sourcePartPath: "ppt/slides/slide1.xml",
      }),
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
