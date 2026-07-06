import { parseTransform } from "../reader/drawing.js";
import { createSidecarIdFactory } from "../reader/raw-node.js";
import { parseShapeTree } from "../reader/shape-tree.js";
import {
  getChild,
  localName,
  navigateOrdered,
  parseXml,
  parseXmlOrdered,
  type XmlNode,
} from "../reader/xml.js";
import type {
  PartPath,
  PptxSourceModel,
  SourceBlipEffects,
  SourceCellBorders,
  SourceEffectList,
  SourceFill,
  SourceImage,
  SourceOutline,
  SourceParagraph,
  SourceParagraphProperties,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceSmartArt,
  SourceStyleReference,
  SourceTable,
  SourceTableCell,
  SourceTextBody,
  SourceTextStyle,
  SourceTheme,
  SourceTransform,
} from "../source/index.js";
import { asEmu } from "../source/index.js";
import { parseComputedChartData } from "./chart-data.js";
import { buildComputedColorScheme, buildEffectiveColorMap, resolveColor } from "./color.js";
import { findPlaceholderMatch } from "./placeholders.js";
import type {
  ComputedBackground,
  ComputedBlipEffects,
  ComputedCellBorders,
  ComputedChartElement,
  ComputedConnectorElement,
  ComputedDiagramDrawing,
  ComputedDiagramDrawingDiagnostic,
  ComputedEffectList,
  ComputedElement,
  ComputedElementLayer,
  ComputedFill,
  ComputedGroupElement,
  ComputedImageElement,
  ComputedOutline,
  ComputedParagraph,
  ComputedRelationship,
  ComputedRunProperties,
  ComputedShapeElement,
  ComputedSlide,
  ComputedSmartArtElement,
  ComputedTableCell,
  ComputedTableElement,
  ComputedTableRow,
  ComputedTextBody,
  CreateComputedViewOptions,
  PptxComputedView,
} from "./pptx-computed-view.js";
import { resolveComputedRelationships } from "./relationships.js";

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const CHART_REL_TYPES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
]);
const DIAGRAM_DATA_REL_TYPES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData",
]);
const DIAGRAM_DRAWING_REL_TYPES: ReadonlySet<string> = new Set([
  "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramDrawing",
]);

const textDecoder = new TextDecoder();

// Preserve the existing renderer-compatibility default for table styles until
// tableStyles.xml is modeled explicitly.
const DEFAULT_TABLE_STYLE_BORDERS: SourceCellBorders = {
  top: {
    width: asEmu(12700),
    fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
  },
  bottom: {
    width: asEmu(12700),
    fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
  },
  left: {
    width: asEmu(12700),
    fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
  },
  right: {
    width: asEmu(12700),
    fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
  },
};

export function createComputedView(
  source: PptxSourceModel,
  options: CreateComputedViewOptions = {},
): PptxComputedView {
  const slidesByPath = new Map(source.slides.map((slide) => [slide.partPath, slide]));
  const selectedSlideNumbers = options.slides !== undefined ? new Set(options.slides) : undefined;

  const slides: ComputedSlide[] = [];
  source.presentation.slidePartPaths.forEach((partPath, index) => {
    const slideNumber = index + 1;
    if (selectedSlideNumbers !== undefined && !selectedSlideNumbers.has(slideNumber)) return;

    const slide = slidesByPath.get(partPath);
    if (slide === undefined) return;
    slides.push(computeSlide(source, slide, slideNumber, options));
  });

  return {
    ...(source.presentation.slideSize !== undefined
      ? { slideSize: { ...source.presentation.slideSize } }
      : {}),
    slides,
  };
}

function computeSlide(
  source: PptxSourceModel,
  slide: SourceSlide,
  slideNumber: number,
  options: CreateComputedViewOptions,
): ComputedSlide {
  const layout = source.slideLayouts.find(
    (candidate) => candidate.partPath === slide.layoutPartPath,
  );
  const master =
    layout !== undefined
      ? source.slideMasters.find((candidate) => candidate.partPath === layout.masterPartPath)
      : undefined;
  const theme =
    master?.themePartPath !== undefined
      ? source.themes.find((candidate) => candidate.partPath === master.themePartPath)
      : undefined;
  const colorMap = buildEffectiveColorMap(
    master?.colorMap,
    layout?.colorMapOverride,
    slide.colorMapOverride,
  );
  const relationships = resolveComputedRelationships(source, slide.partPath);
  const context: ComputeContext = { source, layout, master, theme, colorMap, relationships };
  const colorScheme = buildComputedColorScheme(context);

  const layoutShowMasterShapes = layout?.showMasterShapes ?? true;
  const showMasterShapes = slide.showMasterShapes ?? true;
  const includeMaster =
    options.applyMasterVisibility === false ? true : showMasterShapes && layoutShowMasterShapes;

  return {
    slideNumber,
    partPath: slide.partPath,
    ...(layout !== undefined ? { layoutPartPath: layout.partPath } : {}),
    ...(master !== undefined ? { masterPartPath: master.partPath } : {}),
    ...(theme?.partPath !== undefined ? { themePartPath: theme.partPath } : {}),
    ...(source.presentation.slideSize !== undefined
      ? { slideSize: { ...source.presentation.slideSize } }
      : {}),
    relationships,
    colorMap,
    colorScheme,
    ...(computeBackground(context, slide, layout, master) ?? {}),
    showMasterShapes,
    layoutShowMasterShapes,
    elements: [
      ...(includeMaster && master !== undefined
        ? computeTemplateElements(context, master.shapes, "master", master.partPath)
        : []),
      ...(layout !== undefined
        ? computeTemplateElements(context, layout.shapes, "layout", layout.partPath)
        : []),
      ...computeSlideElements(context, slide.shapes, slide.partPath),
    ],
  };
}

interface ComputeContext {
  readonly source: PptxSourceModel;
  readonly layout?: SourceSlideLayout;
  readonly master?: SourceSlideMaster;
  readonly theme?: SourceTheme;
  readonly colorMap: Readonly<Record<string, string>>;
  readonly relationships: readonly ComputedRelationship[];
  readonly groupFill?: ComputedFill;
}

interface TextStyleChainEntry {
  readonly style: SourceTextStyle;
  readonly includeRunDecorations: boolean;
}

interface RunPropertyDefaults {
  readonly properties?: SourceRunProperties;
  readonly includeDecorations: boolean;
}

type MutableRunProperties = {
  -readonly [K in keyof SourceRunProperties]?: SourceRunProperties[K];
};

function computeBackground(
  context: ComputeContext,
  slide: SourceSlide,
  layout: SourceSlideLayout | undefined,
  master: SourceSlideMaster | undefined,
): { readonly background: ComputedBackground } | undefined {
  const picked =
    slide.background !== undefined
      ? { sourceLayer: "slide" as const, partPath: slide.partPath, background: slide.background }
      : layout?.background !== undefined
        ? {
            sourceLayer: "layout" as const,
            partPath: layout.partPath,
            background: layout.background,
          }
        : master?.background !== undefined
          ? {
              sourceLayer: "master" as const,
              partPath: master.partPath,
              background: master.background,
            }
          : undefined;
  if (picked === undefined) return undefined;

  const { background, sourceLayer } = picked;
  switch (background.kind) {
    case "fill":
      return {
        background: {
          kind: "fill",
          source: background,
          fill: computeFill(context, background.fill, picked.partPath),
          sourceLayer,
        },
      };
    case "styleReference": {
      const color = resolveColor(context, background.color);
      return {
        background: {
          kind: "styleReference",
          source: background,
          index: background.index,
          ...(color !== undefined ? { color } : {}),
          sourceLayer,
        },
      };
    }
    case "raw":
      return { background: { kind: "raw", source: background, sourceLayer } };
  }
}

function computeTemplateElements(
  context: ComputeContext,
  elements: readonly SourceShapeNode[],
  layer: "master" | "layout",
  partPath: PartPath,
): ComputedElement[] {
  return elements
    .filter((element) => !(element.kind === "shape" && element.placeholder !== undefined))
    .map((element) => computeElement(context, element, layer, partPath));
}

function computeSlideElements(
  context: ComputeContext,
  elements: readonly SourceShapeNode[],
  partPath: PartPath,
): ComputedElement[] {
  return elements
    .filter((element) => !(element.kind === "shape" && isEmptyPlaceholder(element)))
    .map((element) => computeElement(context, element, "slide", partPath));
}

function computeElement(
  context: ComputeContext,
  element: SourceShapeNode,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedElement {
  switch (element.kind) {
    case "shape":
      return computeShapeElement(context, element, layer, partPath);
    case "connector":
      return computeConnectorElement(context, element, layer, partPath);
    case "group":
      return computeGroupElement(context, element, layer, partPath);
    case "image":
      return computeImageElement(context, element, layer, partPath);
    case "table":
      return computeTableElement(context, element, layer, partPath);
    case "chart":
      return computeChartElement(context, element, layer, partPath);
    case "smartArt":
      return computeSmartArtElement(context, element, layer, partPath);
    case "raw":
      return { kind: "raw", sourceLayer: layer, sourcePartPath: partPath, sourceNode: element };
  }
}

function computeConnectorElement(
  context: ComputeContext,
  connector: Extract<SourceShapeNode, { kind: "connector" }>,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedConnectorElement {
  const effects =
    connector.effects !== undefined
      ? computeEffectList(context, connector.effects)
      : connector.style?.effectRef !== undefined
        ? resolveEffectReference(context, connector.style.effectRef)
        : undefined;
  return {
    kind: "connector",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: connector,
    ...(connector.transform !== undefined ? { transform: connector.transform } : {}),
    ...(connector.geometry !== undefined ? { geometry: connector.geometry } : {}),
    ...computedOutlineProperty(context, connector.outline, connector.style?.lineRef, partPath),
    ...(effects !== undefined ? { effects } : {}),
  };
}

function computeGroupElement(
  context: ComputeContext,
  group: Extract<SourceShapeNode, { kind: "group" }>,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedGroupElement {
  const fill = group.fill !== undefined ? computeFill(context, group.fill, partPath) : undefined;
  const effects =
    group.effects !== undefined ? computeEffectList(context, group.effects) : undefined;
  const childContext = fill !== undefined ? { ...context, groupFill: fill } : context;
  return {
    kind: "group",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: group,
    ...(group.transform !== undefined ? { transform: group.transform } : {}),
    ...(group.childTransform !== undefined ? { childTransform: group.childTransform } : {}),
    ...(fill !== undefined ? { fill } : {}),
    ...(effects !== undefined ? { effects } : {}),
    children: group.children.map((child) => computeElement(childContext, child, layer, partPath)),
  };
}

function computeShapeElement(
  context: ComputeContext,
  shape: SourceShape,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedShapeElement {
  const match =
    layer === "slide"
      ? findPlaceholderMatch(
          {
            layoutShapes: context.layout?.shapes ?? [],
            masterShapes: context.master?.shapes ?? [],
          },
          shape,
        )
      : undefined;
  const placeholderType =
    shape.placeholder !== undefined ? (shape.placeholder.type ?? "body") : undefined;
  const transform = firstDefined(
    shape.transform,
    match?.layout?.transform,
    match?.master?.transform,
  );
  const geometry = firstDefined(shape.geometry, match?.layout?.geometry, match?.master?.geometry);
  const effects =
    shape.effects !== undefined
      ? computeEffectList(context, shape.effects)
      : shape.style?.effectRef !== undefined
        ? resolveEffectReference(context, shape.style.effectRef)
        : undefined;

  return {
    kind: "shape",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: shape,
    ...(transform !== undefined ? { transform } : {}),
    ...(geometry !== undefined ? { geometry } : {}),
    ...(shape.fill !== undefined
      ? { fill: computeFill(context, shape.fill, partPath) }
      : shape.style?.fillRef !== undefined
        ? { fill: resolveFillReference(context, shape.style.fillRef, partPath) }
        : {}),
    ...computedOutlineProperty(context, shape.outline, shape.style?.lineRef, partPath),
    ...(effects !== undefined ? { effects } : {}),
    ...(shape.textBody !== undefined
      ? {
          textBody: computeTextBody(
            context,
            shape.textBody,
            [match?.layout?.textBody, match?.master?.textBody],
            placeholderType,
          ),
        }
      : {}),
    ...(match !== undefined ? { placeholderMatch: match } : {}),
  };
}

function computeImageElement(
  context: ComputeContext,
  image: SourceImage,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedImageElement {
  const relationship = resolveComputedRelationships(context.source, partPath).find(
    (rel) => rel.id === image.blipRelationshipId && rel.type === IMAGE_REL_TYPE,
  );
  const effects =
    image.effects !== undefined ? computeEffectList(context, image.effects) : undefined;
  const blipEffects =
    image.blipEffects !== undefined ? computeBlipEffects(context, image.blipEffects) : undefined;

  return {
    kind: "image",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: image,
    ...(image.transform !== undefined ? { transform: image.transform } : {}),
    ...(relationship !== undefined ? { relationship } : {}),
    ...(relationship?.media !== undefined ? { media: relationship.media } : {}),
    ...(effects !== undefined ? { effects } : {}),
    ...(blipEffects !== undefined ? { blipEffects } : {}),
  };
}

function computeTableElement(
  context: ComputeContext,
  table: SourceTable,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedTableElement {
  return {
    kind: "table",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: table,
    ...(table.transform !== undefined ? { transform: table.transform } : {}),
    table: {
      columns: table.table.columns.map((column) => ({ ...column })),
      rows: table.table.rows.map((row) => computeTableRow(context, table, row, partPath)),
    },
  };
}

function computeChartElement(
  context: ComputeContext,
  chart: Extract<SourceShapeNode, { kind: "chart" }>,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedChartElement {
  const relationship = context.relationships.find(
    (rel) => rel.id === chart.chartRelationshipId && CHART_REL_TYPES.has(rel.type),
  );
  const chartXml =
    relationship?.targetPartPath !== undefined
      ? readRawPackageText(context.source, relationship.targetPartPath)
      : undefined;
  const chartData = chartXml !== undefined ? parseComputedChartData(chartXml, context) : undefined;
  return {
    kind: "chart",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: chart,
    ...(chart.transform !== undefined ? { transform: chart.transform } : {}),
    ...(relationship !== undefined ? { relationship } : {}),
    ...(chartXml !== undefined ? { chartXml } : {}),
    ...(chartData !== undefined ? { chartData } : {}),
  };
}

function computeSmartArtElement(
  context: ComputeContext,
  smartArt: SourceSmartArt,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedSmartArtElement {
  const dataRelationship = context.relationships.find(
    (rel) => rel.id === smartArt.dataRelationshipId && DIAGRAM_DATA_REL_TYPES.has(rel.type),
  );
  const dataRelationships =
    dataRelationship?.targetPartPath !== undefined
      ? resolveComputedRelationships(context.source, dataRelationship.targetPartPath)
      : [];
  const drawingRelationship = dataRelationships.find((rel) =>
    DIAGRAM_DRAWING_REL_TYPES.has(rel.type),
  );
  const drawingPartPath = drawingRelationship?.targetPartPath;
  const drawingXml =
    drawingPartPath !== undefined ? readRawPackageText(context.source, drawingPartPath) : undefined;
  const drawingRelationships =
    drawingPartPath !== undefined
      ? resolveComputedRelationships(context.source, drawingPartPath)
      : [];
  const diagramDrawing =
    drawingPartPath !== undefined && drawingXml !== undefined
      ? computeDiagramDrawing(context, drawingPartPath, drawingXml, drawingRelationships, layer)
      : undefined;

  return {
    kind: "smartArt",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: smartArt,
    ...(smartArt.transform !== undefined ? { transform: smartArt.transform } : {}),
    ...(dataRelationship !== undefined ? { dataRelationship } : {}),
    ...(drawingRelationship !== undefined ? { drawingRelationship } : {}),
    ...(drawingPartPath !== undefined ? { drawingPartPath } : {}),
    ...(drawingXml !== undefined ? { drawingXml } : {}),
    drawingRelationships,
    media: context.source.packageGraph.media,
    ...(diagramDrawing !== undefined ? { diagramDrawing } : {}),
  };
}

function computeDiagramDrawing(
  context: ComputeContext,
  drawingPartPath: PartPath,
  drawingXml: string,
  drawingRelationships: readonly ComputedRelationship[],
  layer: ComputedElementLayer,
): ComputedDiagramDrawing {
  const rawPart = context.source.packageGraph.rawParts?.find(
    (part) => part.partPath === drawingPartPath,
  );
  const parsed = parseXml(drawingXml);
  const drawing = getChild(parsed, "drawing");
  const spTree = getChild(drawing, "spTree");
  const diagnostics: ComputedDiagramDrawingDiagnostic[] = [];

  if (spTree === undefined) {
    diagnostics.push({
      severity: "warning",
      code: "diagram-drawing-shape-tree-missing",
      message: `diagram drawing part '${drawingPartPath}' has no shape tree`,
      sourcePartPath: drawingPartPath,
    });
  }

  const orderedSpTree =
    spTree !== undefined
      ? navigateOrdered(parseXmlOrdered(drawingXml), ["drawing", "spTree"])
      : undefined;
  const sourceChildren =
    spTree !== undefined
      ? parseShapeTree(
          spTree,
          drawingPartPath,
          createSidecarIdFactory(drawingPartPath),
          orderedSpTree,
        )
      : [];
  const diagramContext: ComputeContext = {
    ...context,
    relationships: drawingRelationships,
  };

  return {
    sourcePartPath: drawingPartPath,
    rawXml: drawingXml,
    ...(rawPart !== undefined ? { rawPart } : {}),
    rawHandle: { partPath: drawingPartPath },
    relationships: drawingRelationships,
    media: context.source.packageGraph.media,
    ...(spTree !== undefined ? { childTransform: parseDiagramChildTransform(spTree) } : {}),
    children: sourceChildren.map((child) =>
      computeElement(diagramContext, child, layer, drawingPartPath),
    ),
    diagnostics,
  };
}

function parseDiagramChildTransform(spTree: XmlNode): SourceTransform | undefined {
  const groupProperties = getChild(spTree, "grpSpPr");
  const fallback = parseTransform(groupProperties);
  const xfrm = getChild(groupProperties, "xfrm");
  const childOff = getChild(xfrm, "chOff");
  const childExt = getChild(xfrm, "chExt");
  const offsetX = numericAttr(childOff, "x") ?? 0;
  const offsetY = numericAttr(childOff, "y") ?? 0;
  const width = numericAttr(childExt, "cx") ?? fallback?.width;
  const height = numericAttr(childExt, "cy") ?? fallback?.height;
  if (width === undefined || height === undefined) return undefined;
  return {
    offsetX: asEmu(offsetX),
    offsetY: asEmu(offsetY),
    width: asEmu(width),
    height: asEmu(height),
  };
}

function computeTableRow(
  context: ComputeContext,
  table: SourceTable,
  row: SourceTable["table"]["rows"][number],
  partPath: PartPath,
): ComputedTableRow {
  return {
    source: row,
    height: row.height,
    cells: row.cells.map((cell) => computeTableCell(context, table, cell, partPath)),
  };
}

function computeTableCell(
  context: ComputeContext,
  table: SourceTable,
  cell: SourceTableCell,
  partPath: PartPath,
): ComputedTableCell {
  const borders =
    cell.borders !== undefined
      ? computeCellBorders(context, cell.borders, partPath)
      : table.table.tableStyleId !== undefined
        ? computeCellBorders(context, DEFAULT_TABLE_STYLE_BORDERS, partPath)
        : undefined;
  return {
    source: cell,
    ...(cell.textBody !== undefined
      ? { textBody: computeTextBody(context, cell.textBody, [], undefined, false) }
      : {}),
    ...(cell.fill !== undefined ? { fill: computeFill(context, cell.fill, partPath) } : {}),
    ...(borders !== undefined ? { borders } : {}),
    gridSpan: cell.gridSpan,
    rowSpan: cell.rowSpan,
    hMerge: cell.hMerge,
    vMerge: cell.vMerge,
  };
}

function computeCellBorders(
  context: ComputeContext,
  borders: SourceCellBorders,
  partPath: PartPath,
): ComputedCellBorders | undefined {
  const computed: ComputedCellBorders = {
    ...(borders.top !== undefined ? { top: computeOutline(context, borders.top, partPath) } : {}),
    ...(borders.bottom !== undefined
      ? { bottom: computeOutline(context, borders.bottom, partPath) }
      : {}),
    ...(borders.left !== undefined
      ? { left: computeOutline(context, borders.left, partPath) }
      : {}),
    ...(borders.right !== undefined
      ? { right: computeOutline(context, borders.right, partPath) }
      : {}),
  };
  return Object.keys(computed).length > 0 ? computed : undefined;
}

function computeFill(context: ComputeContext, fill: SourceFill, partPath: PartPath): ComputedFill {
  switch (fill.kind) {
    case "none":
      return { kind: "none", source: fill };
    case "raw":
      if (localName(fill.raw.node.name) === "grpFill" && context.groupFill !== undefined) {
        return context.groupFill;
      }
      return { kind: "raw", source: fill };
    case "gradient":
      return {
        kind: "gradient",
        source: fill,
        stops: fill.stops.map((stop) => ({
          position: stop.position,
          color: resolveColor(context, stop.color) ?? { hex: "#000000", alpha: 1 },
        })),
        gradientType: fill.gradientType,
        ...(fill.gradientType === "linear" ? { angle: Number(fill.angle) / 60000 } : {}),
        ...(fill.gradientType === "radial" ? { centerX: fill.centerX, centerY: fill.centerY } : {}),
      };
    case "pattern":
      return {
        kind: "pattern",
        source: fill,
        preset: fill.preset,
        foregroundColor: resolveColor(context, fill.foregroundColor) ?? {
          hex: "#000000",
          alpha: 1,
        },
        backgroundColor: resolveColor(context, fill.backgroundColor) ?? {
          hex: "#ffffff",
          alpha: 1,
        },
      };
    case "image": {
      const relationship = resolveComputedRelationships(context.source, partPath).find(
        (rel) => rel.id === fill.blipRelationshipId && rel.type === IMAGE_REL_TYPE,
      );
      return {
        kind: "image",
        source: fill,
        ...(relationship !== undefined ? { relationship } : {}),
        ...(relationship?.media !== undefined ? { media: relationship.media } : {}),
        ...(fill.tile !== undefined ? { tile: fill.tile } : {}),
      };
    }
    case "solid":
      return {
        kind: "solid",
        source: fill,
        color: resolveColor(context, fill.color) ?? { hex: "#000000", alpha: 1 },
      };
  }
}

function computeOutline(
  context: ComputeContext,
  outline: SourceOutline,
  partPath: PartPath,
): ComputedOutline {
  return {
    source: outline,
    ...(outline.width !== undefined ? { width: outline.width } : {}),
    ...(outline.fill !== undefined ? { fill: computeFill(context, outline.fill, partPath) } : {}),
  };
}

function computedOutlineProperty(
  context: ComputeContext,
  outline: SourceOutline | undefined,
  lineRef: SourceStyleReference | undefined,
  partPath: PartPath,
): { readonly outline?: ComputedOutline } {
  const styleOutline =
    lineRef !== undefined ? resolveLineReference(context, lineRef, partPath) : undefined;
  if (outline === undefined) {
    return styleOutline !== undefined ? { outline: styleOutline } : {};
  }

  const computed = mergeComputedOutline(styleOutline, computeOutline(context, outline, partPath));
  return { outline: computed };
}

function mergeComputedOutline(
  base: ComputedOutline | undefined,
  override: ComputedOutline,
): ComputedOutline {
  if (base === undefined) return override;
  const width = override.width ?? base.width;
  const fill = override.fill ?? base.fill;
  return {
    source: { ...base.source, ...override.source },
    ...(width !== undefined ? { width } : {}),
    ...(fill !== undefined ? { fill } : {}),
  };
}

function resolveFillReference(
  context: ComputeContext,
  ref: SourceStyleReference,
  partPath: PartPath,
): ComputedFill | undefined {
  if (ref.index === 0) return undefined;
  // OOXML style references route idx >= 1000 to bgFillStyleLst, with idx=1000
  // resolving to no template.
  const list =
    ref.index >= 1000
      ? context.theme?.formatScheme?.backgroundFillStyles
      : context.theme?.formatScheme?.fillStyles;
  const arrayIndex = ref.index >= 1000 ? ref.index - 1001 : ref.index - 1;
  const template = list?.[arrayIndex];
  if (template === undefined) return undefined;
  const computed = computeFill(context, template, context.theme?.partPath ?? partPath);
  const overrideColor = ref.color !== undefined ? resolveColor(context, ref.color) : undefined;
  if (overrideColor === undefined) return computed;
  if (computed.kind === "solid") {
    return { ...computed, color: overrideColor };
  }
  if (computed.kind === "gradient") {
    return {
      ...computed,
      stops: computed.stops.map((stop) => ({ ...stop, color: overrideColor })),
    };
  }
  return computed;
}

function resolveLineReference(
  context: ComputeContext,
  ref: SourceStyleReference,
  partPath: PartPath,
): ComputedOutline | undefined {
  if (ref.index === 0) return undefined;
  const template = context.theme?.formatScheme?.lineStyles[ref.index - 1];
  if (template === undefined) return undefined;
  const computed = computeOutline(context, template, context.theme?.partPath ?? partPath);
  const overrideSourceColor = ref.color;
  if (overrideSourceColor === undefined) return computed;
  const overrideColor = resolveColor(context, overrideSourceColor);
  if (overrideColor === undefined) return computed;
  return {
    ...computed,
    fill: {
      kind: "solid",
      source: { kind: "solid", color: overrideSourceColor },
      color: overrideColor,
    },
  };
}

function resolveEffectReference(
  context: ComputeContext,
  ref: SourceStyleReference,
): ComputedEffectList | undefined {
  // effectStyleLst is indexed directly by effectRef@idx for renderer
  // compatibility with existing snapshots.
  const template = context.theme?.formatScheme?.effectStyles[ref.index];
  return template !== undefined ? computeEffectList(context, template) : undefined;
}

function computeEffectList(
  context: ComputeContext,
  source: SourceEffectList,
): ComputedEffectList | undefined {
  const outerShadowColor =
    source.outerShadow !== undefined ? resolveColor(context, source.outerShadow.color) : undefined;
  const innerShadowColor =
    source.innerShadow !== undefined ? resolveColor(context, source.innerShadow.color) : undefined;
  const glowColor =
    source.glow !== undefined ? resolveColor(context, source.glow.color) : undefined;
  const computed: ComputedEffectList = {
    source,
    ...(source.outerShadow !== undefined && outerShadowColor !== undefined
      ? {
          outerShadow: {
            blurRadius: source.outerShadow.blurRadius,
            distance: source.outerShadow.distance,
            direction: Number(source.outerShadow.direction) / 60000,
            color: outerShadowColor,
            alignment: source.outerShadow.alignment,
            rotateWithShape: source.outerShadow.rotateWithShape,
          },
        }
      : {}),
    ...(source.innerShadow !== undefined && innerShadowColor !== undefined
      ? {
          innerShadow: {
            blurRadius: source.innerShadow.blurRadius,
            distance: source.innerShadow.distance,
            direction: Number(source.innerShadow.direction) / 60000,
            color: innerShadowColor,
          },
        }
      : {}),
    ...(source.glow !== undefined && glowColor !== undefined
      ? {
          glow: {
            radius: source.glow.radius,
            color: glowColor,
          },
        }
      : {}),
    ...(source.softEdge !== undefined ? { softEdge: source.softEdge } : {}),
  };
  return computed.outerShadow !== undefined ||
    computed.innerShadow !== undefined ||
    computed.glow !== undefined ||
    computed.softEdge !== undefined
    ? computed
    : undefined;
}

function computeBlipEffects(
  context: ComputeContext,
  source: SourceBlipEffects,
): ComputedBlipEffects | undefined {
  const duotoneColor1 =
    source.duotone !== undefined ? resolveColor(context, source.duotone.color1) : undefined;
  const duotoneColor2 =
    source.duotone !== undefined ? resolveColor(context, source.duotone.color2) : undefined;
  const clrChangeFrom =
    source.clrChange !== undefined ? resolveColor(context, source.clrChange.from) : undefined;
  const clrChangeTo =
    source.clrChange !== undefined ? resolveColor(context, source.clrChange.to) : undefined;
  const computed: ComputedBlipEffects = {
    source,
    grayscale: source.grayscale,
    ...(source.biLevel !== undefined ? { biLevel: source.biLevel } : {}),
    ...(source.blur !== undefined ? { blur: source.blur } : {}),
    ...(source.lum !== undefined ? { lum: source.lum } : {}),
    ...(duotoneColor1 !== undefined && duotoneColor2 !== undefined
      ? { duotone: { color1: duotoneColor1, color2: duotoneColor2 } }
      : {}),
    ...(clrChangeFrom !== undefined && clrChangeTo !== undefined
      ? { clrChange: { from: clrChangeFrom, to: clrChangeTo } }
      : {}),
  };
  return computed.grayscale ||
    computed.biLevel !== undefined ||
    computed.blur !== undefined ||
    computed.lum !== undefined ||
    computed.duotone !== undefined ||
    computed.clrChange !== undefined
    ? computed
    : undefined;
}

function computeTextBody(
  context: ComputeContext,
  textBody: SourceTextBody,
  inheritedBodies: readonly (SourceTextBody | undefined)[] = [],
  placeholderType?: string,
  includeInheritedStyleChain = true,
): ComputedTextBody {
  const styleChain = buildTextStyleChain(
    context,
    textBody,
    inheritedBodies,
    placeholderType,
    includeInheritedStyleChain,
  );
  // Keep placeholder bodyPr inheritance disabled until the document path
  // intentionally owns that behavior.
  const properties = mergeTextBodyProperties(undefined, textBody.properties);
  return {
    ...(properties !== undefined ? { properties } : {}),
    paragraphs: textBody.paragraphs.map((paragraph) =>
      computeParagraph(context, paragraph, styleChain),
    ),
  };
}

function mergeTextBodyProperties(
  inherited: SourceTextBody["properties"] | undefined,
  local: SourceTextBody["properties"] | undefined,
): SourceTextBody["properties"] | undefined {
  const merged = { ...inherited, ...local };
  if (Object.keys(merged).length === 0) return undefined;
  if (merged.autoFit === "normAutofit") {
    return {
      ...merged,
      fontScale: merged.fontScale ?? 1,
      lnSpcReduction: merged.lnSpcReduction ?? 0,
    };
  }
  if (merged.autoFit === "spAutofit" || merged.autoFit === "noAutofit") {
    return {
      ...merged,
      fontScale: 1,
      lnSpcReduction: 0,
    };
  }
  return merged;
}

function computeParagraph(
  context: ComputeContext,
  paragraph: SourceParagraph,
  styleChain: readonly TextStyleChainEntry[],
): ComputedParagraph {
  const level = paragraph.properties?.level ?? 0;
  const styleLevelProperties = styleChain.map((entry) => ({
    properties: textStyleLevelProperties(entry.style, level),
    includeRunDecorations: entry.includeRunDecorations,
  }));
  const properties = computeParagraphProperties(
    context,
    paragraph.properties,
    styleLevelProperties.map((entry) => entry.properties),
  );
  return {
    ...(properties !== undefined ? { properties } : {}),
    runs: paragraph.runs.map((run) => {
      const runProperties = mergeRunProperties(context, run.properties, [
        {
          properties: paragraph.properties?.defaultRunProperties,
          includeDecorations: true,
        },
        ...styleLevelProperties.map((entry) => ({
          properties: entry.properties?.defaultRunProperties,
          includeDecorations: entry.includeRunDecorations,
        })),
      ]);
      return {
        text: run.text,
        ...(runProperties !== undefined ? { properties: runProperties } : {}),
      };
    }),
  };
}

function buildTextStyleChain(
  context: ComputeContext,
  textBody: SourceTextBody,
  inheritedBodies: readonly (SourceTextBody | undefined)[],
  placeholderType: string | undefined,
  includeInheritedStyleChain: boolean,
): TextStyleChainEntry[] {
  return [
    ...(textBody.listStyle !== undefined
      ? [{ style: textBody.listStyle, includeRunDecorations: true }]
      : []),
    ...(includeInheritedStyleChain
      ? [
          ...inheritedBodies.flatMap((body) =>
            body?.listStyle !== undefined
              ? [{ style: body.listStyle, includeRunDecorations: false }]
              : [],
          ),
          ...styleEntry(getTxStyleForPlaceholder(context.master?.txStyles, placeholderType)),
          ...styleEntry(context.source.presentation.defaultTextStyle),
        ]
      : []),
  ];
}

function getTxStyleForPlaceholder(
  txStyles: NonNullable<ComputeContext["master"]>["txStyles"] | undefined,
  placeholderType: string | undefined,
): SourceTextStyle | undefined {
  if (txStyles === undefined) return undefined;
  if (placeholderType === undefined) return txStyles.otherStyle;
  if (placeholderType === "title" || placeholderType === "ctrTitle") return txStyles.titleStyle;
  if (placeholderType === "body" || placeholderType === "subTitle" || placeholderType === "obj") {
    return txStyles.bodyStyle;
  }
  return txStyles.otherStyle;
}

function textStyleLevelProperties(
  style: SourceTextStyle,
  level: number,
): SourceParagraphProperties | undefined {
  return style.levels[level] ?? style.defaultParagraph;
}

function computeParagraphProperties(
  context: ComputeContext,
  local: SourceParagraphProperties | undefined,
  inherited: readonly (SourceParagraphProperties | undefined)[],
): ComputedParagraph["properties"] {
  const resolvedBulletColor = firstDefined(
    local?.bulletColor !== undefined ? resolveColor(context, local.bulletColor) : undefined,
    ...inherited.map((properties) =>
      properties?.bulletColor !== undefined
        ? resolveColor(context, properties.bulletColor)
        : undefined,
    ),
  );
  const merged: ComputedParagraph["properties"] = {
    ...(local?.align !== undefined
      ? { align: local.align }
      : { align: pickInheritedValue(inherited, "align") ?? "left" }),
    ...(local?.level !== undefined ? { level: local.level } : { level: 0 }),
    ...(local?.lineSpacing !== undefined ? { lineSpacing: local.lineSpacing } : {}),
    ...(local?.spaceBefore !== undefined ? { spaceBefore: local.spaceBefore } : {}),
    ...(local?.spaceAfter !== undefined ? { spaceAfter: local.spaceAfter } : {}),
    ...(local?.marginLeft !== undefined
      ? { marginLeft: local.marginLeft }
      : pickInherited(inherited, "marginLeft")),
    ...(local?.indent !== undefined
      ? { indent: local.indent }
      : pickInherited(inherited, "indent")),
    ...(local?.bullet !== undefined
      ? { bullet: local.bullet }
      : pickInherited(inherited, "bullet")),
    ...(local?.bulletFont !== undefined
      ? { bulletFont: local.bulletFont }
      : pickInherited(inherited, "bulletFont")),
    ...(resolvedBulletColor !== undefined ? { bulletColor: resolvedBulletColor } : {}),
    ...(local?.bulletSizePct !== undefined
      ? { bulletSizePct: local.bulletSizePct }
      : pickInherited(inherited, "bulletSizePct")),
    ...(local?.tabStops !== undefined ? { tabStops: local.tabStops } : {}),
  };
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function mergeRunProperties(
  context: ComputeContext,
  local: SourceRunProperties | undefined,
  inherited: readonly RunPropertyDefaults[],
): ComputedRunProperties | undefined {
  const merged: {
    bold?: SourceRunProperties["bold"];
    italic?: SourceRunProperties["italic"];
    underline?: SourceRunProperties["underline"];
    strikethrough?: SourceRunProperties["strikethrough"];
    baseline?: SourceRunProperties["baseline"];
    fontSize?: SourceRunProperties["fontSize"];
    typeface?: SourceRunProperties["typeface"];
    typefaceEa?: SourceRunProperties["typefaceEa"];
    typefaceCs?: SourceRunProperties["typefaceCs"];
    color?: SourceRunProperties["color"];
  } = { ...local };
  for (const defaults of inherited) {
    if (defaults.properties === undefined) continue;
    mergeDefaultRunProperties(merged, defaults.properties, defaults.includeDecorations);
  }
  if (Object.keys(merged).length === 0) return undefined;
  const color = merged.color !== undefined ? resolveColor(context, merged.color) : undefined;
  const withoutColor = resolveThemeRunFonts(context, omitColor(merged));
  return {
    ...withoutColor,
    ...(color !== undefined ? { color } : {}),
  };
}

function styleEntry(style: SourceTextStyle | undefined): TextStyleChainEntry[] {
  return style !== undefined ? [{ style, includeRunDecorations: false }] : [];
}

function mergeDefaultRunProperties(
  target: MutableRunProperties,
  defaults: SourceRunProperties,
  includeDecorations: boolean,
): void {
  if (includeDecorations) {
    target.bold ??= defaults.bold;
    target.italic ??= defaults.italic;
    target.underline ??= defaults.underline;
    target.strikethrough ??= defaults.strikethrough;
    target.baseline ??= defaults.baseline;
  }
  target.fontSize ??= defaults.fontSize;
  target.typeface ??= defaults.typeface;
  target.typefaceEa ??= defaults.typefaceEa;
  target.typefaceCs ??= defaults.typefaceCs;
  target.color ??= defaults.color;
}

function pickInherited<K extends keyof SourceParagraphProperties>(
  inherited: readonly (SourceParagraphProperties | undefined)[],
  key: K,
): Partial<Pick<SourceParagraphProperties, K>> {
  const value = pickInheritedValue(inherited, key);
  const picked: Partial<Pick<SourceParagraphProperties, K>> = {};
  if (value !== undefined) {
    picked[key] = value;
  }
  return picked;
}

function pickInheritedValue<K extends keyof SourceParagraphProperties>(
  inherited: readonly (SourceParagraphProperties | undefined)[],
  key: K,
): SourceParagraphProperties[K] | undefined {
  const value = inherited.find((properties) => properties?.[key] !== undefined)?.[key];
  return value;
}

function resolveThemeRunFonts(
  context: ComputeContext,
  properties: Omit<SourceRunProperties, "color">,
): Omit<SourceRunProperties, "color"> {
  return {
    ...properties,
    ...(properties.typeface !== undefined
      ? { typeface: resolveThemeTypeface(context, properties.typeface) }
      : {}),
    ...(properties.typefaceEa !== undefined
      ? { typefaceEa: resolveThemeTypeface(context, properties.typefaceEa) }
      : {}),
    ...(properties.typefaceCs !== undefined
      ? { typefaceCs: resolveThemeTypeface(context, properties.typefaceCs) }
      : {}),
  };
}

function resolveThemeTypeface(context: ComputeContext, typeface: string): string {
  const scheme = context.theme?.fontScheme;
  if (scheme === undefined) return typeface;
  switch (typeface) {
    case "+mj-lt":
      return scheme.majorLatin ?? typeface;
    case "+mn-lt":
      return scheme.minorLatin ?? typeface;
    case "+mj-ea":
      return scheme.majorEastAsian ?? scheme.majorJapanese ?? typeface;
    case "+mn-ea":
      return scheme.minorEastAsian ?? scheme.minorJapanese ?? typeface;
    case "+mj-cs":
      return scheme.majorComplexScript ?? typeface;
    case "+mn-cs":
      return scheme.minorComplexScript ?? typeface;
    default:
      return typeface;
  }
}

function isEmptyPlaceholder(shape: SourceShape): boolean {
  if (shape.placeholder === undefined) return false;
  const paragraphs = shape.textBody?.paragraphs;
  if (paragraphs === undefined || paragraphs.length === 0) return true;
  return !paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0));
}

function readRawPackageText(source: PptxSourceModel, partPath: PartPath): string | undefined {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart?.kind === "binary") return textDecoder.decode(rawPart.bytes);
  return undefined;
}

function numericAttr(node: XmlNode | undefined, attrName: string): number | undefined {
  if (node === undefined) return undefined;
  const value = node[`@_${attrName}`];
  if (value === undefined || value === null) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function omitColor(properties: SourceRunProperties): Omit<SourceRunProperties, "color"> {
  const withoutColor = { ...properties };
  delete withoutColor.color;
  return withoutColor;
}

function firstDefined<T>(...values: readonly (T | undefined)[]): T | undefined {
  return values.find((value) => value !== undefined);
}
