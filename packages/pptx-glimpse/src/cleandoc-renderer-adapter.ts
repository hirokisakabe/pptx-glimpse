import type {
  Background,
  BodyProperties,
  Fill,
  Geometry,
  ImageElement,
  Outline,
  Paragraph,
  ParagraphProperties,
  ResolvedColor,
  RunProperties,
  Slide,
  SlideElement,
  SlideSize,
  TextBody,
  TextRun,
  Transform,
} from "pptx-glimpse-renderer";
import { asEmu, asHundredthPt, uint8ArrayToBase64 } from "pptx-glimpse-renderer";

import type {
  CleanDocComputedView,
  ComputedBackground,
  ComputedColor,
  ComputedElement,
  ComputedFill,
  ComputedImageElement,
  ComputedOutline,
  ComputedParagraph,
  ComputedRunProperties,
  ComputedShapeElement,
  ComputedSlide,
  ComputedSlideSize,
  ComputedTextBody,
} from "../../pptx-glimpse-document/src/experimental.js";

interface RendererAdapterResult {
  readonly slideSize?: SlideSize;
  readonly slides: readonly Slide[];
  readonly diagnostics: readonly RendererAdapterDiagnostic[];
}

export interface RendererAdapterDiagnostic {
  readonly severity: "warning";
  readonly code:
    | "cleandoc-adapter.missing-transform"
    | "cleandoc-adapter.raw-element-skipped"
    | "cleandoc-adapter.raw-background-ignored"
    | "cleandoc-adapter.raw-fill-ignored"
    | "cleandoc-adapter.unresolved-image-skipped";
  readonly message: string;
  readonly slideNumber?: number;
  readonly sourcePartPath?: string;
}

type DiagnosticSink = RendererAdapterDiagnostic[];

const ZERO_TRANSFORM: Transform = {
  offsetX: asEmu(0),
  offsetY: asEmu(0),
  extentWidth: asEmu(0),
  extentHeight: asEmu(0),
  rotation: 0,
  flipH: false,
  flipV: false,
};

export function adaptComputedViewToRendererModel(
  computed: CleanDocComputedView,
): RendererAdapterResult {
  const diagnostics: DiagnosticSink = [];

  return {
    ...(computed.slideSize !== undefined ? { slideSize: adaptSlideSize(computed.slideSize) } : {}),
    slides: computed.slides.map((slide) => adaptSlide(slide, diagnostics)),
    diagnostics,
  };
}

function adaptSlide(slide: ComputedSlide, diagnostics: DiagnosticSink): Slide {
  return {
    slideNumber: slide.slideNumber,
    background: adaptBackground(slide.background, slide, diagnostics),
    elements: slide.elements.flatMap((element) => adaptElement(element, slide, diagnostics)),
    showMasterSp: slide.showMasterShapes,
  };
}

function adaptSlideSize(slideSize: ComputedSlideSize): SlideSize {
  return {
    width: toRendererEmu(slideSize.width),
    height: toRendererEmu(slideSize.height),
  };
}

function adaptBackground(
  background: ComputedBackground | undefined,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): Background | null {
  if (background === undefined) return null;
  if (background.kind === "fill") {
    return { fill: adaptFill(background.fill, slide, diagnostics) };
  }
  if (background.kind === "styleReference") {
    return {
      fill:
        background.color !== undefined
          ? { type: "solid", color: adaptColor(background.color) }
          : null,
    };
  }

  diagnostics.push({
    severity: "warning",
    code: "cleandoc-adapter.raw-background-ignored",
    message: "Raw CleanDoc background is not supported by the renderer adapter.",
    slideNumber: slide.slideNumber,
  });
  return null;
}

function adaptElement(
  element: ComputedElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): SlideElement[] {
  if (element.kind === "shape") return [adaptShape(element, slide, diagnostics)];
  if (element.kind === "image") {
    const image = adaptImage(element, slide, diagnostics);
    return image === undefined ? [] : [image];
  }

  diagnostics.push({
    severity: "warning",
    code: "cleandoc-adapter.raw-element-skipped",
    message: "Raw CleanDoc shape tree node is outside the renderer adapter subset.",
    slideNumber: slide.slideNumber,
    sourcePartPath: element.sourcePartPath,
  });
  return [];
}

function adaptShape(
  shape: ComputedShapeElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): SlideElement {
  return {
    type: "shape",
    transform: adaptTransform(shape.transform, slide, diagnostics, shape.sourcePartPath),
    geometry: adaptGeometry(shape.geometry),
    fill: shape.fill !== undefined ? adaptFill(shape.fill, slide, diagnostics) : null,
    outline: shape.outline !== undefined ? adaptOutline(shape.outline, slide, diagnostics) : null,
    textBody: shape.textBody !== undefined ? adaptTextBody(shape.textBody) : null,
    effects: null,
    ...(shape.sourceNode.placeholder?.type !== undefined
      ? { placeholderType: shape.sourceNode.placeholder.type }
      : {}),
    ...(shape.sourceNode.placeholder?.index !== undefined
      ? { placeholderIdx: shape.sourceNode.placeholder.index }
      : {}),
    ...(shape.sourceNode.name !== undefined ? { altText: shape.sourceNode.name } : {}),
  };
}

function adaptImage(
  image: ComputedImageElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): ImageElement | undefined {
  if (image.media === undefined) {
    diagnostics.push({
      severity: "warning",
      code: "cleandoc-adapter.unresolved-image-skipped",
      message: "CleanDoc image element has no resolved media payload.",
      slideNumber: slide.slideNumber,
      sourcePartPath: image.sourcePartPath,
    });
    return undefined;
  }

  return {
    type: "image",
    transform: adaptTransform(image.transform, slide, diagnostics, image.sourcePartPath),
    imageData: uint8ArrayToBase64(image.media.bytes),
    mimeType: image.media.contentType,
    effects: null,
    blipEffects: null,
    srcRect: image.sourceNode.crop
      ? {
          left: ooxmlPercentToRatio(image.sourceNode.crop.left),
          top: ooxmlPercentToRatio(image.sourceNode.crop.top),
          right: ooxmlPercentToRatio(image.sourceNode.crop.right),
          bottom: ooxmlPercentToRatio(image.sourceNode.crop.bottom),
        }
      : null,
    ...(image.sourceNode.name !== undefined ? { altText: image.sourceNode.name } : {}),
    stretch: null,
    tile: null,
  };
}

function adaptTransform(
  transform: ComputedShapeElement["transform"] | undefined,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
  sourcePartPath: string,
): Transform {
  if (transform === undefined) {
    diagnostics.push({
      severity: "warning",
      code: "cleandoc-adapter.missing-transform",
      message: "CleanDoc element has no computed transform; using a zero-size fallback.",
      slideNumber: slide.slideNumber,
      sourcePartPath,
    });
    return ZERO_TRANSFORM;
  }

  return {
    offsetX: toRendererEmu(transform.offsetX),
    offsetY: toRendererEmu(transform.offsetY),
    extentWidth: toRendererEmu(transform.width),
    extentHeight: toRendererEmu(transform.height),
    rotation: Number(transform.rotation ?? 0) / 60000,
    flipH: transform.flipHorizontal ?? false,
    flipV: transform.flipVertical ?? false,
  };
}

function adaptGeometry(geometry: ComputedShapeElement["geometry"] | undefined): Geometry {
  return {
    type: "preset",
    preset: geometry?.preset ?? "rect",
    adjustValues: {},
  };
}

function adaptFill(
  fill: ComputedFill,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): Fill | null {
  if (fill.kind === "none") return { type: "none" };
  if (fill.kind === "solid") {
    return { type: "solid", color: adaptColor(fill.color) };
  }

  diagnostics.push({
    severity: "warning",
    code: "cleandoc-adapter.raw-fill-ignored",
    message: "Raw CleanDoc fill is outside the renderer adapter subset.",
    slideNumber: slide.slideNumber,
  });
  return null;
}

function adaptOutline(
  outline: ComputedOutline,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): Outline | null {
  const fill = outline.fill !== undefined ? adaptFill(outline.fill, slide, diagnostics) : null;
  if (fill?.type === "none") return null;

  return {
    width: toRendererEmu(outline.width ?? 12700),
    fill: fill?.type === "solid" || fill?.type === "gradient" ? fill : null,
    dashStyle: "solid",
    headEnd: null,
    tailEnd: null,
  };
}

function adaptTextBody(textBody: ComputedTextBody): TextBody {
  return {
    bodyProperties: adaptBodyProperties(textBody.properties),
    paragraphs: textBody.paragraphs.map(adaptParagraph),
  };
}

function adaptBodyProperties(properties: ComputedTextBody["properties"]): BodyProperties {
  return {
    anchor: properties?.anchor !== undefined ? adaptAnchor(properties.anchor) : "t",
    marginLeft: toRendererEmu(properties?.marginLeft ?? 91440),
    marginRight: toRendererEmu(properties?.marginRight ?? 91440),
    marginTop: toRendererEmu(properties?.marginTop ?? 45720),
    marginBottom: toRendererEmu(properties?.marginBottom ?? 45720),
    wrap: "square",
    autoFit: "noAutofit",
    fontScale: 1,
    lnSpcReduction: 0,
    numCol: 1,
    vert: "horz",
  };
}

function adaptAnchor(anchor: "top" | "middle" | "bottom"): BodyProperties["anchor"] {
  if (anchor === "middle") return "ctr";
  if (anchor === "bottom") return "b";
  return "t";
}

function adaptParagraph(paragraph: ComputedParagraph): Paragraph {
  return {
    properties: adaptParagraphProperties(paragraph.properties),
    runs: paragraph.runs.map(
      (run): TextRun => ({
        text: run.text,
        properties: adaptRunProperties(run.properties),
      }),
    ),
  };
}

function adaptParagraphProperties(
  properties: ComputedParagraph["properties"],
): ParagraphProperties {
  return {
    alignment: properties?.align !== undefined ? adaptAlignment(properties.align) : null,
    lineSpacing:
      properties?.lineSpacingPts !== undefined
        ? { type: "pts", value: asHundredthPt(Number(properties.lineSpacingPts)) }
        : null,
    spaceBefore: { type: "pts", value: asHundredthPt(0) },
    spaceAfter: { type: "pts", value: asHundredthPt(0) },
    level: properties?.level ?? 0,
    bullet: adaptBullet(properties?.bullet),
    bulletFont: properties?.bulletFont ?? null,
    bulletColor: null,
    bulletSizePct: null,
    marginLeft: null,
    indent: null,
    tabStops: [],
  };
}

function adaptBullet(
  bullet: NonNullable<ComputedParagraph["properties"]>["bullet"] | undefined,
): ParagraphProperties["bullet"] {
  if (bullet === undefined) return null;
  return bullet;
}

function adaptAlignment(
  alignment: NonNullable<ComputedParagraph["properties"]>["align"],
): ParagraphProperties["alignment"] {
  if (alignment === "center") return "ctr";
  if (alignment === "right") return "r";
  if (alignment === "justify") return "just";
  return "l";
}

function adaptRunProperties(properties: ComputedRunProperties | undefined): RunProperties {
  return {
    fontSize: properties?.fontSize !== undefined ? toRendererPt(properties.fontSize) : null,
    fontFamily: properties?.typeface ?? null,
    fontFamilyEa: null,
    fontFamilyCs: null,
    bold: properties?.bold ?? false,
    italic: properties?.italic ?? false,
    underline: properties?.underline ?? false,
    strikethrough: false,
    color: properties?.color !== undefined ? adaptColor(properties.color) : null,
    baseline: 0,
    hyperlink: null,
    outline: null,
  };
}

function adaptColor(color: ComputedColor): ResolvedColor {
  return {
    hex: color.hex,
    alpha: color.alpha,
  };
}

function ooxmlPercentToRatio(value: number | undefined): number {
  return Number(value ?? 0) / 100000;
}

function toRendererEmu(value: number): ReturnType<typeof asEmu> {
  return asEmu(Number(value));
}

function toRendererPt(value: number): NonNullable<RunProperties["fontSize"]> {
  return Number(value) as NonNullable<RunProperties["fontSize"]>;
}
