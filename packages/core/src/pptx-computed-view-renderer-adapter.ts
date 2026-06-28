/**
 * PptxComputedView から renderer model への core-owned adapter。
 *
 * `@pptx-glimpse/document` は source/computed document semantics を所有し、
 * `@pptx-glimpse/renderer` は SVG/PNG generation 向けの display-oriented model
 * を所有する。この adapter は両者を統合せず、computed view の provenance-rich
 * effective values を renderer の render-ready contract へ写す境界である。
 *
 * Adapter 側では renderer-specific defaults、missing transform fallback、
 * `null` fill/outline/background convention、base64 media payload、chart XML to
 * renderer ChartData conversion、SmartArt diagram drawing fallback、raw
 * element/fill/background warning policy を扱う。これらは writer/editor/round-trip
 * に必要な PptxSourceModel semantics ではないため document package に戻さない。
 *
 * Font discovery/fallback、text measurement/wrapping、text-to-path、SVG/PNG output
 * choices は renderer/core 側に残す。ComputedSlide/ComputedElement の part path、
 * source layer、source node、relationships、theme context は diagnostics と
 * fallback conversion の入力として使い、renderer model へは必要な render contract
 * だけを渡す。
 */

import type {
  ComputedBackground,
  ComputedBlipEffects,
  ComputedChartElement,
  ComputedColor,
  ComputedConnectorElement,
  ComputedEffectList,
  ComputedElement,
  ComputedFill,
  ComputedGroupElement,
  ComputedImageElement,
  ComputedOutline,
  ComputedParagraph,
  ComputedRunProperties,
  ComputedShapeElement,
  ComputedSlide,
  ComputedSlideSize,
  ComputedSmartArtElement,
  ComputedTableElement,
  ComputedTextBody,
  PptxComputedView,
} from "@pptx-glimpse/document";
import type {
  Background,
  BlipEffects,
  BodyProperties,
  ChartElement,
  ColorMap,
  ColorScheme,
  ConnectorElement,
  EffectList,
  Fill,
  Geometry,
  GroupElement,
  ImageElement,
  Outline,
  Paragraph,
  ParagraphProperties,
  ResolvedColor,
  RunProperties,
  Slide,
  SlideElement,
  SlideSize,
  SpacingValue,
  TableCell,
  TableElement,
  TextBody,
  TextRun,
  Transform,
} from "@pptx-glimpse/renderer";
import { asEmu, asHundredthPt, uint8ArrayToBase64 } from "@pptx-glimpse/renderer";

import { ColorResolver } from "./color/color-resolver.js";
import { convertChartXmlToRendererChartData } from "./renderer-chart-data-converter.js";
import { unsafeBrandAssertion } from "./unsafe-type-assertion.js";

interface RendererAdapterResult {
  readonly slideSize?: SlideSize;
  readonly slides: readonly Slide[];
  readonly diagnostics: readonly RendererAdapterDiagnostic[];
}

export interface RendererAdapterDiagnostic {
  readonly severity: "warning";
  readonly code:
    | "pptx-computed-view-adapter.missing-transform"
    | "pptx-computed-view-adapter.raw-element-skipped"
    | "pptx-computed-view-adapter.raw-background-ignored"
    | "pptx-computed-view-adapter.raw-fill-ignored"
    | "pptx-computed-view-adapter.unresolved-chart-skipped"
    | "pptx-computed-view-adapter.unresolved-smartart-skipped"
    | "pptx-computed-view-adapter.unresolved-image-skipped";
  readonly message: string;
  readonly slideNumber?: number;
  readonly sourcePartPath?: string;
}

type DiagnosticSink = RendererAdapterDiagnostic[];
type RendererAdapterDiagnosticCode = RendererAdapterDiagnostic["code"];

const ZERO_TRANSFORM: Transform = {
  offsetX: asEmu(0),
  offsetY: asEmu(0),
  extentWidth: asEmu(0),
  extentHeight: asEmu(0),
  rotation: 0,
  flipH: false,
  flipV: false,
};

const DEFAULT_COLOR_SCHEME: ColorScheme = {
  dk1: "#000000",
  lt1: "#ffffff",
  dk2: "#44546a",
  lt2: "#e7e6e6",
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
  hlink: "#0563c1",
  folHlink: "#954f72",
};

const DEFAULT_COLOR_MAP: ColorMap = {
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
};

export function adaptComputedViewToRendererModel(
  computed: PptxComputedView,
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
  switch (background.kind) {
    case "fill":
      return { fill: adaptFill(background.fill, slide, diagnostics) };
    case "styleReference":
      return {
        fill:
          background.color !== undefined
            ? { type: "solid", color: adaptColor(background.color) }
            : null,
      };
    case "raw":
      pushAdapterWarning(
        diagnostics,
        "pptx-computed-view-adapter.raw-background-ignored",
        "Raw PptxSourceModel background is not supported by the renderer adapter.",
        slide,
      );
      return null;
    default:
      return assertNever(background);
  }
}

function adaptElement(
  element: ComputedElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): SlideElement[] {
  switch (element.kind) {
    case "shape":
      return [adaptShape(element, slide, diagnostics)];
    case "connector":
      return [adaptConnector(element, slide, diagnostics)];
    case "group":
      return [adaptGroup(element, slide, diagnostics)];
    case "image": {
      const image = adaptImage(element, slide, diagnostics);
      return image === undefined ? [] : [image];
    }
    case "table":
      return [adaptTable(element, slide, diagnostics)];
    case "chart": {
      const chart = adaptChart(element, slide, diagnostics);
      return chart === undefined ? [] : [chart];
    }
    case "smartArt": {
      const smartArt = adaptSmartArt(element, slide, diagnostics);
      return smartArt === undefined ? [] : [smartArt];
    }
    case "raw":
      pushAdapterWarning(
        diagnostics,
        "pptx-computed-view-adapter.raw-element-skipped",
        "Raw PptxSourceModel shape tree node is outside the renderer adapter subset.",
        slide,
        element.sourcePartPath,
      );
      return [];
    default:
      return assertNever(element);
  }
}

function adaptConnector(
  connector: ComputedConnectorElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): ConnectorElement {
  return {
    type: "connector",
    transform: adaptTransform(connector.transform, slide, diagnostics, connector.sourcePartPath),
    geometry: adaptGeometry(connector.geometry),
    outline:
      connector.outline !== undefined ? adaptOutline(connector.outline, slide, diagnostics) : null,
    effects: connector.effects !== undefined ? adaptEffects(connector.effects) : null,
    ...(connector.sourceNode.name !== undefined ? { altText: connector.sourceNode.name } : {}),
  };
}

function adaptGroup(
  group: ComputedGroupElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): GroupElement {
  const transform = adaptTransform(group.transform, slide, diagnostics, group.sourcePartPath);
  return {
    type: "group",
    transform,
    childTransform:
      group.childTransform !== undefined
        ? adaptTransform(group.childTransform, slide, diagnostics, group.sourcePartPath)
        : {
            offsetX: asEmu(0),
            offsetY: asEmu(0),
            extentWidth: transform.extentWidth,
            extentHeight: transform.extentHeight,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
    children: group.children.flatMap((child) => adaptElement(child, slide, diagnostics)),
    effects: group.effects !== undefined ? adaptEffects(group.effects) : null,
    ...(group.sourceNode.name !== undefined ? { altText: group.sourceNode.name } : {}),
  };
}

function adaptChart(
  chart: ComputedChartElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): ChartElement | undefined {
  if (chart.chartXml === undefined) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.unresolved-chart-skipped",
      "PptxSourceModel chart element has no resolved chart XML.",
      slide,
      chart.sourcePartPath,
    );
    return undefined;
  }

  const chartData = convertChartXmlToRendererChartData(chart.chartXml, createColorResolver(slide));
  if (chartData === null) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.unresolved-chart-skipped",
      "PptxSourceModel chart XML could not be parsed into the renderer chart model.",
      slide,
      chart.sourcePartPath,
    );
    return undefined;
  }

  return {
    type: "chart",
    transform: adaptTransform(chart.transform, slide, diagnostics, chart.sourcePartPath),
    chart: chartData,
  };
}

function adaptSmartArt(
  smartArt: ComputedSmartArtElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): GroupElement | undefined {
  if (smartArt.diagramDrawing === undefined) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.unresolved-smartart-skipped",
      "PptxSourceModel SmartArt element has no computed diagram drawing view.",
      slide,
      smartArt.sourcePartPath,
    );
    return undefined;
  }

  if (smartArt.diagramDrawing.diagnostics.length > 0) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.unresolved-smartart-skipped",
      smartArt.diagramDrawing.diagnostics[0]?.message ??
        "PptxSourceModel SmartArt diagram drawing could not be computed.",
      slide,
      smartArt.diagramDrawing.sourcePartPath,
    );
    return undefined;
  }

  const children = smartArt.diagramDrawing.children.flatMap((child) =>
    adaptElement(child, slide, diagnostics),
  );
  if (children.length === 0) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.unresolved-smartart-skipped",
      "PptxSourceModel SmartArt computed diagram drawing has no renderer-supported children.",
      slide,
      smartArt.diagramDrawing.sourcePartPath,
    );
    return undefined;
  }

  const groupTransform = adaptTransform(
    smartArt.transform,
    slide,
    diagnostics,
    smartArt.sourcePartPath,
  );
  const childTransform =
    smartArt.diagramDrawing.childTransform !== undefined
      ? adaptTransform(
          smartArt.diagramDrawing.childTransform,
          slide,
          diagnostics,
          smartArt.diagramDrawing.sourcePartPath,
        )
      : {
          offsetX: asEmu(0),
          offsetY: asEmu(0),
          extentWidth: groupTransform.extentWidth,
          extentHeight: groupTransform.extentHeight,
          rotation: 0,
          flipH: false,
          flipV: false,
        };

  return {
    type: "group",
    transform: groupTransform,
    childTransform,
    children,
    effects: null,
    ...(smartArt.sourceNode.name !== undefined ? { altText: smartArt.sourceNode.name } : {}),
  };
}

function adaptTable(
  table: ComputedTableElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): TableElement {
  return {
    type: "table",
    transform: adaptTransform(table.transform, slide, diagnostics, table.sourcePartPath),
    table: {
      columns: table.table.columns.map((column) => ({ width: toRendererEmu(column.width) })),
      rows: table.table.rows.map((row) => ({
        height: toRendererEmu(row.height),
        cells: row.cells.map((cell): TableCell => {
          return {
            textBody: cell.textBody !== undefined ? adaptTextBody(cell.textBody) : null,
            fill: cell.fill !== undefined ? adaptFill(cell.fill, slide, diagnostics) : null,
            borders:
              cell.borders !== undefined
                ? {
                    top: adaptTableBorder(cell.borders.top, slide, diagnostics),
                    bottom: adaptTableBorder(cell.borders.bottom, slide, diagnostics),
                    left: adaptTableBorder(cell.borders.left, slide, diagnostics),
                    right: adaptTableBorder(cell.borders.right, slide, diagnostics),
                  }
                : null,
            gridSpan: cell.gridSpan,
            rowSpan: cell.rowSpan,
            hMerge: cell.hMerge,
            vMerge: cell.vMerge,
          };
        }),
      })),
    },
  };
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
    effects: shape.effects !== undefined ? adaptEffects(shape.effects) : null,
    ...(shape.sourceNode.placeholder !== undefined
      ? { placeholderType: shape.sourceNode.placeholder.type ?? "body" }
      : {}),
    ...(shape.sourceNode.placeholder?.index !== undefined
      ? { placeholderIdx: shape.sourceNode.placeholder.index }
      : {}),
    ...(shape.sourceNode.name !== undefined ? { altText: shape.sourceNode.name } : {}),
  };
}

function adaptEffects(effects: ComputedEffectList): EffectList {
  return {
    outerShadow:
      effects.outerShadow !== undefined
        ? {
            blurRadius: toRendererEmu(effects.outerShadow.blurRadius),
            distance: toRendererEmu(effects.outerShadow.distance),
            direction: effects.outerShadow.direction,
            color: adaptColor(effects.outerShadow.color),
            alignment: effects.outerShadow.alignment,
            rotateWithShape: effects.outerShadow.rotateWithShape,
          }
        : null,
    innerShadow:
      effects.innerShadow !== undefined
        ? {
            blurRadius: toRendererEmu(effects.innerShadow.blurRadius),
            distance: toRendererEmu(effects.innerShadow.distance),
            direction: effects.innerShadow.direction,
            color: adaptColor(effects.innerShadow.color),
          }
        : null,
    glow:
      effects.glow !== undefined
        ? {
            radius: toRendererEmu(effects.glow.radius),
            color: adaptColor(effects.glow.color),
          }
        : null,
    softEdge:
      effects.softEdge !== undefined
        ? {
            radius: toRendererEmu(effects.softEdge.radius),
          }
        : null,
  };
}

function adaptImage(
  image: ComputedImageElement,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): ImageElement | undefined {
  if (image.media === undefined) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.unresolved-image-skipped",
      "PptxSourceModel image element has no resolved media payload.",
      slide,
      image.sourcePartPath,
    );
    return undefined;
  }

  return {
    type: "image",
    transform: adaptTransform(image.transform, slide, diagnostics, image.sourcePartPath),
    imageData: uint8ArrayToBase64(image.media.bytes),
    mimeType: normalizeImageMimeType(image.media.contentType),
    effects: image.effects !== undefined ? adaptEffects(image.effects) : null,
    blipEffects: image.blipEffects !== undefined ? adaptBlipEffects(image.blipEffects) : null,
    srcRect: image.sourceNode.crop
      ? {
          left: ooxmlPercentToRatio(image.sourceNode.crop.left),
          top: ooxmlPercentToRatio(image.sourceNode.crop.top),
          right: ooxmlPercentToRatio(image.sourceNode.crop.right),
          bottom: ooxmlPercentToRatio(image.sourceNode.crop.bottom),
        }
      : null,
    ...(image.sourceNode.name !== undefined ? { altText: image.sourceNode.name } : {}),
    stretch: image.sourceNode.stretch ?? null,
    tile:
      image.sourceNode.tile !== undefined
        ? {
            tx: toRendererEmu(image.sourceNode.tile.tx),
            ty: toRendererEmu(image.sourceNode.tile.ty),
            sx: image.sourceNode.tile.sx,
            sy: image.sourceNode.tile.sy,
            flip: image.sourceNode.tile.flip,
            align: image.sourceNode.tile.align,
          }
        : null,
  };
}

function adaptBlipEffects(effects: ComputedBlipEffects): BlipEffects {
  return {
    grayscale: effects.grayscale,
    biLevel: effects.biLevel ?? null,
    blur:
      effects.blur !== undefined
        ? { radius: toRendererEmu(effects.blur.radius), grow: effects.blur.grow }
        : null,
    lum: effects.lum ?? null,
    duotone:
      effects.duotone !== undefined
        ? {
            color1: adaptColor(effects.duotone.color1),
            color2: adaptColor(effects.duotone.color2),
          }
        : null,
    clrChange:
      effects.clrChange !== undefined
        ? {
            clrFrom: adaptColor(effects.clrChange.from),
            clrTo: adaptColor(effects.clrChange.to),
          }
        : null,
  };
}

function adaptTransform(
  transform: ComputedShapeElement["transform"] | undefined,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
  sourcePartPath: string,
): Transform {
  if (transform === undefined) {
    pushAdapterWarning(
      diagnostics,
      "pptx-computed-view-adapter.missing-transform",
      "PptxSourceModel element has no computed transform; using a zero-size fallback.",
      slide,
      sourcePartPath,
    );
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
  if (geometry === undefined) {
    return {
      type: "preset",
      preset: "rect",
      adjustValues: {},
    };
  }
  if ("paths" in geometry) {
    return { type: "custom", paths: [...geometry.paths] };
  }
  return {
    type: "preset",
    preset: geometry.preset,
    adjustValues: geometry.adjustValues ?? {},
  };
}

function adaptFill(
  fill: ComputedFill,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): Fill | null {
  switch (fill.kind) {
    case "none":
      return { type: "none" };
    case "solid":
      return { type: "solid", color: adaptColor(fill.color) };
    case "gradient":
      return {
        type: "gradient",
        stops: fill.stops.map((stop) => ({
          position: stop.position,
          color: adaptColor(stop.color),
        })),
        angle: fill.angle ?? 0,
        gradientType: fill.gradientType,
        ...(fill.centerX !== undefined ? { centerX: fill.centerX } : {}),
        ...(fill.centerY !== undefined ? { centerY: fill.centerY } : {}),
      };
    case "pattern":
      return {
        type: "pattern",
        preset: fill.preset,
        foregroundColor: adaptColor(fill.foregroundColor),
        backgroundColor: adaptColor(fill.backgroundColor),
      };
    case "image":
      if (fill.media !== undefined) {
        return {
          type: "image",
          imageData: uint8ArrayToBase64(fill.media.bytes),
          mimeType: normalizeImageMimeType(fill.media.contentType),
          tile:
            fill.tile !== undefined
              ? {
                  tx: toRendererEmu(fill.tile.tx),
                  ty: toRendererEmu(fill.tile.ty),
                  sx: fill.tile.sx,
                  sy: fill.tile.sy,
                  flip: fill.tile.flip,
                  align: fill.tile.align,
                }
              : null,
        };
      }
      pushAdapterWarning(
        diagnostics,
        "pptx-computed-view-adapter.raw-fill-ignored",
        "Raw PptxSourceModel fill is outside the renderer adapter subset.",
        slide,
      );
      return null;
    case "raw":
      pushAdapterWarning(
        diagnostics,
        "pptx-computed-view-adapter.raw-fill-ignored",
        "Raw PptxSourceModel fill is outside the renderer adapter subset.",
        slide,
      );
      return null;
    default:
      return assertNever(fill);
  }
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
    dashStyle: outline.source.dashStyle ?? "solid",
    ...(outline.source.customDash !== undefined
      ? { customDash: [...outline.source.customDash] }
      : {}),
    ...(outline.source.lineCap !== undefined ? { lineCap: outline.source.lineCap } : {}),
    ...(outline.source.lineJoin !== undefined ? { lineJoin: outline.source.lineJoin } : {}),
    headEnd: outline.source.headEnd ?? null,
    tailEnd: outline.source.tailEnd ?? null,
  };
}

function adaptTableBorder(
  outline: ComputedOutline | undefined,
  slide: ComputedSlide,
  diagnostics: DiagnosticSink,
): Outline | null {
  if (outline === undefined) return null;
  const border = adaptOutline(outline, slide, diagnostics);
  if (border === null) return null;
  return {
    ...border,
    fill: border.fill ?? { type: "solid", color: { hex: "#000000", alpha: 1 } },
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
    wrap: properties?.wrap ?? "square",
    autoFit: properties?.autoFit ?? "noAutofit",
    fontScale: properties?.fontScale ?? 1,
    lnSpcReduction: properties?.lnSpcReduction ?? 0,
    numCol: properties?.numCol ?? 1,
    vert: properties?.vert ?? "horz",
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
      properties?.lineSpacing !== undefined ? adaptSpacing(properties.lineSpacing) : null,
    spaceBefore:
      properties?.spaceBefore !== undefined
        ? adaptSpacing(properties.spaceBefore)
        : { type: "pts", value: asHundredthPt(0) },
    spaceAfter:
      properties?.spaceAfter !== undefined
        ? adaptSpacing(properties.spaceAfter)
        : { type: "pts", value: asHundredthPt(0) },
    level: properties?.level ?? 0,
    bullet: properties?.bullet ?? null,
    bulletFont: properties?.bulletFont ?? null,
    bulletColor: properties?.bulletColor !== undefined ? adaptColor(properties.bulletColor) : null,
    bulletSizePct: properties?.bulletSizePct ?? null,
    marginLeft: properties?.marginLeft !== undefined ? toRendererEmu(properties.marginLeft) : null,
    indent: properties?.indent !== undefined ? toRendererEmu(properties.indent) : null,
    tabStops:
      properties?.tabStops?.map((tab) => ({
        position: toRendererEmu(tab.position),
        alignment: tab.alignment,
      })) ?? [],
  };
}

function adaptSpacing(
  spacing: NonNullable<NonNullable<ComputedParagraph["properties"]>["lineSpacing"]>,
): SpacingValue {
  if (spacing.type === "pts") return { type: "pts", value: asHundredthPt(Number(spacing.value)) };
  return { type: "pct", value: spacing.value };
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
    fontFamilyEa: properties?.typefaceEa ?? null,
    fontFamilyCs: properties?.typefaceCs ?? null,
    bold: properties?.bold ?? false,
    italic: properties?.italic ?? false,
    underline: properties?.underline ?? false,
    strikethrough: properties?.strikethrough ?? false,
    color: properties?.color !== undefined ? adaptColor(properties.color) : null,
    baseline: properties?.baseline ?? 0,
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

function createColorResolver(slide: ComputedSlide): ColorResolver {
  return new ColorResolver(
    { ...DEFAULT_COLOR_SCHEME, ...slide.colorScheme },
    {
      ...DEFAULT_COLOR_MAP,
      ...slide.colorMap,
    },
  );
}

function ooxmlPercentToRatio(value: number | undefined): number {
  return Number(value ?? 0) / 100000;
}

function toRendererEmu(value: number): ReturnType<typeof asEmu> {
  return asEmu(Number(value));
}

function toRendererPt(value: number): NonNullable<RunProperties["fontSize"]> {
  return unsafeBrandAssertion<NonNullable<RunProperties["fontSize"]>>(Number(value));
}

function normalizeImageMimeType(contentType: string): string {
  if (contentType === "image/x-emf") return "image/emf";
  if (contentType === "image/x-wmf") return "image/wmf";
  return contentType;
}

function pushAdapterWarning(
  diagnostics: DiagnosticSink,
  code: RendererAdapterDiagnosticCode,
  message: string,
  slide: ComputedSlide,
  sourcePartPath?: string,
): void {
  diagnostics.push({
    severity: "warning",
    code,
    message,
    slideNumber: slide.slideNumber,
    ...(sourcePartPath !== undefined ? { sourcePartPath } : {}),
  });
}

function assertNever(value: never): never {
  throw new Error(`Unexpected computed view union member: ${JSON.stringify(value)}`);
}
