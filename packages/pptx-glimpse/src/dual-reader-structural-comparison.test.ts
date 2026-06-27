import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { createComputedView, readPptx } from "@pptx-glimpse/document/experimental";
import type {
  Background,
  BodyProperties,
  Fill,
  ImageElement,
  Outline,
  Paragraph,
  RunProperties,
  ShapeElement,
  Slide,
  SlideElement,
  SlideSize,
  TextBody,
  Transform,
} from "pptx-glimpse-renderer";
import { describe, expect, it } from "vitest";

import { adaptComputedViewToRendererModel } from "./cleandoc-renderer-adapter.js";
import { buildEffectiveSlideElements } from "./parser-path-oracle.js";
import { parsePptxData, parseSlideWithLayout } from "./pptx-data-parser.js";

const CASES = [
  { fixtureName: "real-product-page.pptx", includeRunProperties: true },
  // real-basic-theme uses inherited placeholder/theme text styles that the
  // document path does not expose yet. Flip this to true when that support lands.
  { fixtureName: "real-basic-theme.pptx", includeRunProperties: false },
] as const;

const COMPARISON_SCOPE = [
  "slide size",
  "slide background fill",
  "effective element ordering after template placeholder filtering",
  "simple shape transform / preset geometry / placeholder metadata",
  "simple shape solid fill / outline colors resolved through theme data",
  "plain text runs and basic run styling",
  "embedded raster image transform / media type / payload presence / crop rectangle",
] as const;

const OUT_OF_SCOPE_FIELDS = [
  "alt text / source node names",
  "shape geometry adjustment handles",
  "effects, shadows, gradients, pattern fills, and picture fills on shapes",
  "raw graphicFrame content, groups, and other raw shape tree nodes",
  "complex paragraph features such as bullets, numbering, tab stops, and hyperlinks",
  "inherited placeholder/theme text styling when it is not yet exposed by the document path",
  "renderer-only text defaults such as East Asian / complex script fallback fields",
] as const;

describe("dual-reader structural comparison", () => {
  it("comparison scope is intentionally limited to the CleanDoc PoC supported subset", () => {
    expect(COMPARISON_SCOPE).toMatchInlineSnapshot(`
      [
        "slide size",
        "slide background fill",
        "effective element ordering after template placeholder filtering",
        "simple shape transform / preset geometry / placeholder metadata",
        "simple shape solid fill / outline colors resolved through theme data",
        "plain text runs and basic run styling",
        "embedded raster image transform / media type / payload presence / crop rectangle",
      ]
    `);
    expect(OUT_OF_SCOPE_FIELDS).toMatchInlineSnapshot(`
      [
        "alt text / source node names",
        "shape geometry adjustment handles",
        "effects, shadows, gradients, pattern fills, and picture fills on shapes",
        "raw graphicFrame content, groups, and other raw shape tree nodes",
        "complex paragraph features such as bullets, numbering, tab stops, and hyperlinks",
        "inherited placeholder/theme text styling when it is not yet exposed by the document path",
        "renderer-only text defaults such as East Asian / complex script fallback fields",
      ]
    `);
  });

  it.each(CASES)(
    "$fixtureName matches between current parser and CleanDoc document path for supported fields",
    ({ fixtureName, includeRunProperties }) => {
      const input = readFixture(fixtureName);
      const options = { includeRunProperties };
      const current = normalizePresentation(buildCurrentParserSlides(input), options);
      const document = buildDocumentPathSlides(input, options);

      for (const diagnostic of document.diagnostics) {
        expect(diagnostic).toMatchObject({
          severity: "warning",
          code: "cleandoc-adapter.raw-element-skipped",
        });
      }
      expect(document.presentation).toEqual(current);
    },
  );
});

interface PresentationSlides {
  readonly slideSize?: SlideSize;
  readonly slides: readonly Slide[];
}

interface DocumentPathSlides {
  readonly presentation: ComparablePresentation;
  readonly diagnostics: readonly { readonly severity: string; readonly code: string }[];
}

interface ComparablePresentation {
  readonly slideSize?: ComparableSlideSize;
  readonly slides: readonly ComparableSlide[];
}

interface ComparableSlideSize {
  readonly width: number;
  readonly height: number;
}

interface ComparableSlide {
  readonly slideNumber: number;
  readonly background: ComparableBackground;
  readonly elements: readonly ComparableElement[];
}

type ComparableBackground = { readonly fill: ComparableFill } | null;

type ComparableElement = ComparableShape | ComparableImage;

interface ComparableShape {
  readonly type: "shape";
  readonly transform: ComparableTransform;
  readonly geometry: ComparableGeometry;
  readonly placeholderType?: string;
  readonly placeholderIdx?: number;
  readonly fill: ComparableFill;
  readonly outline: ComparableOutline;
  readonly textBody: ComparableTextBody | null;
}

interface ComparableImage {
  readonly type: "image";
  readonly transform: ComparableTransform;
  readonly mimeType: string;
  readonly imageData: string;
  readonly srcRect: ImageElement["srcRect"];
}

interface ComparableTransform {
  readonly offsetX: number;
  readonly offsetY: number;
  readonly extentWidth: number;
  readonly extentHeight: number;
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

interface ComparableGeometry {
  readonly type: "preset";
  readonly preset: string;
}

type ComparableFill =
  | { readonly type: "none" }
  | { readonly type: "solid"; readonly color: ComparableColor }
  | null;

type ComparableOutline = { readonly width: number; readonly fill: ComparableFill } | null;

interface ComparableColor {
  readonly hex: string;
  readonly alpha: number;
}

interface ComparableTextBody {
  readonly bodyProperties: ComparableBodyProperties;
  readonly paragraphs: readonly ComparableParagraph[];
}

interface ComparableBodyProperties {
  readonly anchor: BodyProperties["anchor"];
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
}

interface ComparableParagraph {
  readonly runs: readonly ComparableRun[];
}

interface ComparableRun {
  readonly text: string;
  readonly properties: ComparableRunProperties;
}

interface ComparableRunProperties {
  readonly fontSize?: number | null;
  readonly fontFamily?: string | null;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly color?: ComparableColor | null;
}

interface NormalizeOptions {
  readonly includeRunProperties: boolean;
}

function readFixture(name: string): Buffer {
  return readFileSync(fileURLToPath(new URL(`../../../shared-fixtures/${name}`, import.meta.url)));
}

function buildCurrentParserSlides(input: Buffer): PresentationSlides {
  const data = parsePptxData(input);
  const slides: Slide[] = [];

  for (const { slideNumber, path } of data.slidePaths) {
    const parsed = parseSlideWithLayout(slideNumber, path, data);
    if (parsed === null) continue;

    slides.push({
      ...parsed.slide,
      elements: buildEffectiveSlideElements(parsed),
    });
  }

  return { slideSize: data.presInfo.slideSize, slides };
}

function buildDocumentPathSlides(input: Buffer, options: NormalizeOptions): DocumentPathSlides {
  const source = readPptx(input);
  const computed = createComputedView(source);
  const adapted = adaptComputedViewToRendererModel(computed);

  return {
    presentation: normalizePresentation(adapted, options),
    diagnostics: adapted.diagnostics,
  };
}

function normalizePresentation(
  presentation: PresentationSlides,
  options: NormalizeOptions,
): ComparablePresentation {
  return {
    ...(presentation.slideSize !== undefined
      ? { slideSize: normalizeSlideSize(presentation.slideSize) }
      : {}),
    slides: presentation.slides.map((slide) => ({
      slideNumber: slide.slideNumber,
      background: normalizeBackground(slide.background),
      elements: slide.elements.flatMap((element) => normalizeElement(element, options)),
    })),
  };
}

function normalizeSlideSize(slideSize: SlideSize): ComparableSlideSize {
  return {
    width: Number(slideSize.width),
    height: Number(slideSize.height),
  };
}

function normalizeBackground(background: Background | null): ComparableBackground {
  if (background === null) return null;
  return { fill: normalizeFill(background.fill) };
}

function normalizeElement(element: SlideElement, options: NormalizeOptions): ComparableElement[] {
  if (element.type === "shape") return [normalizeShape(element, options)];
  if (element.type === "image") return [normalizeImage(element)];
  return [];
}

function normalizeShape(shape: ShapeElement, options: NormalizeOptions): ComparableShape {
  return {
    type: "shape",
    transform: normalizeTransform(shape.transform),
    geometry: {
      type: "preset",
      preset: shape.geometry.type === "preset" ? shape.geometry.preset : "unsupported",
    },
    ...(shape.placeholderType !== undefined ? { placeholderType: shape.placeholderType } : {}),
    ...(shape.placeholderIdx !== undefined ? { placeholderIdx: shape.placeholderIdx } : {}),
    fill: normalizeFill(shape.fill),
    outline: normalizeOutline(shape.outline),
    textBody: normalizeTextBody(shape.textBody, options),
  };
}

function normalizeImage(image: ImageElement): ComparableImage {
  return {
    type: "image",
    transform: normalizeTransform(image.transform),
    mimeType: image.mimeType,
    imageData: image.imageData,
    srcRect: image.srcRect,
  };
}

function normalizeTransform(transform: Transform): ComparableTransform {
  return {
    offsetX: Number(transform.offsetX),
    offsetY: Number(transform.offsetY),
    extentWidth: Number(transform.extentWidth),
    extentHeight: Number(transform.extentHeight),
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
  };
}

function normalizeFill(fill: Fill | null): ComparableFill {
  if (fill === null) return null;
  if (fill.type === "none") return { type: "none" };
  if (fill.type === "solid") return { type: "solid", color: normalizeColor(fill.color) };
  return null;
}

function normalizeOutline(outline: Outline | null): ComparableOutline {
  if (outline === null) return null;
  return {
    width: Number(outline.width),
    fill: normalizeFill(outline.fill),
  };
}

function normalizeTextBody(
  textBody: TextBody | null,
  options: NormalizeOptions,
): ComparableTextBody | null {
  if (textBody === null) return null;
  return {
    bodyProperties: normalizeBodyProperties(textBody.bodyProperties),
    paragraphs: textBody.paragraphs.map((paragraph) => normalizeParagraph(paragraph, options)),
  };
}

function normalizeBodyProperties(properties: BodyProperties): ComparableBodyProperties {
  return {
    anchor: properties.anchor,
    marginLeft: Number(properties.marginLeft),
    marginRight: Number(properties.marginRight),
    marginTop: Number(properties.marginTop),
    marginBottom: Number(properties.marginBottom),
  };
}

function normalizeParagraph(paragraph: Paragraph, options: NormalizeOptions): ComparableParagraph {
  return {
    runs: paragraph.runs.map((run) => ({
      text: run.text,
      properties: normalizeRunProperties(run.properties, options),
    })),
  };
}

function normalizeRunProperties(
  properties: RunProperties,
  options: NormalizeOptions,
): ComparableRunProperties {
  if (!options.includeRunProperties) return {};
  return {
    fontSize: properties.fontSize === null ? null : Number(properties.fontSize),
    fontFamily: properties.fontFamily,
    bold: properties.bold,
    italic: properties.italic,
    underline: properties.underline,
    color: properties.color === null ? null : normalizeColor(properties.color),
  };
}

function normalizeColor(color: { readonly hex: string; readonly alpha: number }): ComparableColor {
  return {
    hex: color.hex.toLowerCase(),
    alpha: color.alpha,
  };
}
