import type {
  PartPath,
  PptxSourceModel,
  SourceBlipEffects,
  SourceCellBorders,
  SourceColor,
  SourceColorMap,
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
} from "../source/index.js";
import { asEmu, asPartPath } from "../source/index.js";
import type {
  ComputedBackground,
  ComputedBlipEffects,
  ComputedCellBorders,
  ComputedChartElement,
  ComputedColor,
  ComputedConnectorElement,
  ComputedEffectList,
  ComputedElement,
  ComputedElementLayer,
  ComputedFill,
  ComputedGroupElement,
  ComputedImageElement,
  ComputedOutline,
  ComputedParagraph,
  ComputedPlaceholderMatch,
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

const DEFAULT_COLOR_MAP: Readonly<Record<string, string>> = {
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

const FALLBACK_SCHEME_COLORS: Readonly<Record<string, string>> = {
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

// Current parser path treats any tableStyleId as black 1pt cell borders when
// a cell has no inline border definition. Keep this compatibility approximation
// until tableStyles.xml is modeled explicitly.
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
  const relationships = resolveRelationships(source, slide.partPath);
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

function buildEffectiveColorMap(
  master?: SourceColorMap,
  layoutOverride?: SourceColorMap,
  slideOverride?: SourceColorMap,
): Readonly<Record<string, string>> {
  return {
    ...DEFAULT_COLOR_MAP,
    ...master?.mapping,
    ...layoutOverride?.mapping,
    ...slideOverride?.mapping,
  };
}

function buildComputedColorScheme(context: ComputeContext): Readonly<Record<string, string>> {
  const colors: Record<string, string> = { ...FALLBACK_SCHEME_COLORS };
  for (const [name, color] of Object.entries(context.theme?.colorScheme?.colors ?? {})) {
    const resolved = resolveColor(context, color);
    if (resolved !== undefined) colors[name] = resolved.hex;
  }
  return colors;
}

function resolveRelationships(
  source: PptxSourceModel,
  sourcePartPath: PartPath,
): ComputedRelationship[] {
  const rels =
    source.packageGraph.relationships.find((entry) => entry.sourcePartPath === sourcePartPath)
      ?.relationships ?? [];

  return rels.map((relationship) => {
    const external = relationship.targetMode === "External";
    const target = external
      ? relationship.target
      : resolveRelationshipTarget(sourcePartPath, relationship.target);
    const targetPartPath = external ? undefined : asPartPath(target);
    const media =
      targetPartPath !== undefined
        ? source.packageGraph.media.find((part) => part.partPath === targetPartPath)
        : undefined;

    return {
      id: relationship.id,
      type: relationship.type,
      source: relationship,
      target,
      ...(relationship.targetMode !== undefined ? { targetMode: relationship.targetMode } : {}),
      ...(targetPartPath !== undefined ? { targetPartPath } : {}),
      ...(media !== undefined ? { media } : {}),
    };
  });
}

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
  if (background.kind === "fill") {
    return {
      background: {
        kind: "fill",
        source: background,
        fill: computeFill(context, background.fill, picked.partPath),
        sourceLayer,
      },
    };
  }
  if (background.kind === "styleReference") {
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
  return { background: { kind: "raw", source: background, sourceLayer } };
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
  if (element.kind === "shape") return computeShapeElement(context, element, layer, partPath);
  if (element.kind === "connector") {
    return computeConnectorElement(context, element, layer, partPath);
  }
  if (element.kind === "group") return computeGroupElement(context, element, layer, partPath);
  if (element.kind === "image") return computeImageElement(context, element, layer, partPath);
  if (element.kind === "table") return computeTableElement(context, element, layer, partPath);
  if (element.kind === "chart") return computeChartElement(context, element, layer, partPath);
  if (element.kind === "smartArt") {
    return computeSmartArtElement(context, element, layer, partPath);
  }
  return { kind: "raw", sourceLayer: layer, sourcePartPath: partPath, sourceNode: element };
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
    ...(connector.outline !== undefined
      ? { outline: computeOutline(context, connector.outline, partPath) }
      : connector.style?.lineRef !== undefined
        ? { outline: resolveLineReference(context, connector.style.lineRef, partPath) }
        : {}),
    ...(effects !== undefined ? { effects } : {}),
  };
}

function computeGroupElement(
  context: ComputeContext,
  group: Extract<SourceShapeNode, { kind: "group" }>,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedGroupElement {
  const effects =
    group.effects !== undefined ? computeEffectList(context, group.effects) : undefined;
  return {
    kind: "group",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: group,
    ...(group.transform !== undefined ? { transform: group.transform } : {}),
    ...(group.childTransform !== undefined ? { childTransform: group.childTransform } : {}),
    ...(effects !== undefined ? { effects } : {}),
    children: group.children.map((child) => computeElement(context, child, layer, partPath)),
  };
}

function computeShapeElement(
  context: ComputeContext,
  shape: SourceShape,
  layer: ComputedElementLayer,
  partPath: PartPath,
): ComputedShapeElement {
  const match = layer === "slide" ? findPlaceholderMatch(context, shape) : undefined;
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
    ...(shape.outline !== undefined
      ? { outline: computeOutline(context, shape.outline, partPath) }
      : shape.style?.lineRef !== undefined
        ? { outline: resolveLineReference(context, shape.style.lineRef, partPath) }
        : {}),
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
  const relationship = resolveRelationships(context.source, partPath).find(
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
  return {
    kind: "chart",
    sourceLayer: layer,
    sourcePartPath: partPath,
    sourceNode: chart,
    ...(chart.transform !== undefined ? { transform: chart.transform } : {}),
    ...(relationship !== undefined ? { relationship } : {}),
    ...(chartXml !== undefined ? { chartXml } : {}),
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
      ? resolveRelationships(context.source, dataRelationship.targetPartPath)
      : [];
  const drawingRelationship = dataRelationships.find((rel) =>
    DIAGRAM_DRAWING_REL_TYPES.has(rel.type),
  );
  const drawingPartPath = drawingRelationship?.targetPartPath;
  const drawingXml =
    drawingPartPath !== undefined ? readRawPackageText(context.source, drawingPartPath) : undefined;

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
    drawingRelationships:
      drawingPartPath !== undefined ? resolveRelationships(context.source, drawingPartPath) : [],
    media: context.source.packageGraph.media,
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
  if (fill.kind === "none") return { kind: "none", source: fill };
  if (fill.kind === "raw") return { kind: "raw", source: fill };
  if (fill.kind === "gradient") {
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
  }
  if (fill.kind === "pattern") {
    return {
      kind: "pattern",
      source: fill,
      preset: fill.preset,
      foregroundColor: resolveColor(context, fill.foregroundColor) ?? { hex: "#000000", alpha: 1 },
      backgroundColor: resolveColor(context, fill.backgroundColor) ?? { hex: "#ffffff", alpha: 1 },
    };
  }
  if (fill.kind === "image") {
    const relationship = resolveRelationships(context.source, partPath).find(
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
  return {
    kind: "solid",
    source: fill,
    color: resolveColor(context, fill.color) ?? { hex: "#000000", alpha: 1 },
  };
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

function resolveFillReference(
  context: ComputeContext,
  ref: SourceStyleReference,
  partPath: PartPath,
): ComputedFill | undefined {
  if (ref.index === 0) return undefined;
  // Match the current parser oracle in style-reference-resolver.ts: idx >= 1000
  // is routed to bgFillStyleLst, with idx=1000 resolving to no template.
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
  // Match the current parser oracle in style-reference-resolver.ts, where
  // effectStyleLst is indexed directly by effectRef@idx.
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
  // Current parser path does not inherit placeholder bodyPr; keep computed output
  // aligned until the document path intentionally owns that behavior.
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
): Pick<SourceParagraphProperties, K> | Record<string, never> {
  const value = pickInheritedValue(inherited, key);
  return value !== undefined ? ({ [key]: value } as Pick<SourceParagraphProperties, K>) : {};
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

function findPlaceholderMatch(
  context: ComputeContext,
  shape: SourceShape,
): ComputedPlaceholderMatch | undefined {
  if (shape.placeholder === undefined) return undefined;
  const type = shape.placeholder.type ?? "body";
  const index = shape.placeholder.index;
  const layout = findMatchingPlaceholder(type, index, context.layout?.shapes ?? []);
  const master = findMatchingPlaceholder(type, index, context.master?.shapes ?? []);
  if (layout === undefined && master === undefined) return undefined;
  return {
    ...(layout !== undefined ? { layout } : {}),
    ...(master !== undefined ? { master } : {}),
  };
}

function findMatchingPlaceholder(
  type: string,
  index: number | undefined,
  shapes: readonly SourceShapeNode[],
): SourceShape | undefined {
  const placeholders = shapes.filter(
    (shape): shape is SourceShape => shape.kind === "shape" && shape.placeholder !== undefined,
  );

  if (index !== undefined) {
    const byIndexWithTransform = placeholders.find(
      (shape) => shape.placeholder?.index === index && shape.transform !== undefined,
    );
    if (byIndexWithTransform !== undefined) return byIndexWithTransform;
    const byIndex = placeholders.find((shape) => shape.placeholder?.index === index);
    if (byIndex !== undefined) return byIndex;
  }

  const byTypeWithTransform = placeholders.find(
    (shape) => (shape.placeholder?.type ?? "body") === type && shape.transform !== undefined,
  );
  if (byTypeWithTransform !== undefined) return byTypeWithTransform;

  const fallbackType = getPlaceholderFallbackType(type);
  if (fallbackType !== undefined) {
    const byFallbackWithTransform = placeholders.find(
      (shape) =>
        (shape.placeholder?.type ?? "body") === fallbackType && shape.transform !== undefined,
    );
    if (byFallbackWithTransform !== undefined) return byFallbackWithTransform;
  }

  const byType = placeholders.find((shape) => (shape.placeholder?.type ?? "body") === type);
  if (byType !== undefined) return byType;

  if (fallbackType !== undefined) {
    return placeholders.find((shape) => (shape.placeholder?.type ?? "body") === fallbackType);
  }

  return undefined;
}

function getPlaceholderFallbackType(type: string): string | undefined {
  if (type === "ctrTitle") return "title";
  if (type === "subTitle") return "body";
  return undefined;
}

function isEmptyPlaceholder(shape: SourceShape): boolean {
  if (shape.placeholder === undefined) return false;
  const paragraphs = shape.textBody?.paragraphs;
  if (paragraphs === undefined || paragraphs.length === 0) return true;
  return !paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0));
}

function resolveColor(
  context: ComputeContext,
  color: SourceColor,
  visited: ReadonlySet<string> = new Set(),
): ComputedColor | undefined {
  let hex: string | undefined;
  if (color.kind === "srgb") {
    hex = normalizeHex(color.hex);
  } else if (color.kind === "system") {
    hex = normalizeHex(color.lastColor ?? "000000");
  } else {
    const mappedName = context.colorMap[color.scheme] ?? color.scheme;
    if (visited.has(mappedName)) return { hex: "#000000", alpha: 1 };
    const schemeColor = context.theme?.colorScheme?.colors[mappedName];
    hex =
      schemeColor !== undefined
        ? resolveColor(context, schemeColor, new Set([...visited, mappedName]))?.hex
        : FALLBACK_SCHEME_COLORS[mappedName];
  }
  if (hex === undefined) return undefined;

  let alpha = 1;
  for (const transform of color.transforms ?? []) {
    if (transform.kind === "lumMod") {
      const lumOff = color.transforms?.find((candidate) => candidate.kind === "lumOff");
      hex = applyLuminance(
        hex,
        Number(transform.value) / 100000,
        Number(lumOff?.value ?? 0) / 100000,
      );
    } else if (transform.kind === "lumOff") {
      if (!color.transforms?.some((candidate) => candidate.kind === "lumMod")) {
        hex = applyLuminance(hex, 1, Number(transform.value) / 100000);
      }
    } else if (transform.kind === "tint") {
      hex = applyTint(hex, Number(transform.value) / 100000);
    } else if (transform.kind === "shade") {
      hex = applyShade(hex, Number(transform.value) / 100000);
    } else if (transform.kind === "alpha") {
      alpha = Number(transform.value) / 100000;
    }
  }

  return { hex, alpha };
}

function resolveRelationshipTarget(sourcePartPath: string, target: string): string {
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(target)) return target;

  let combined: string;
  if (target.startsWith("/")) {
    combined = target.slice(1);
  } else {
    const slash = sourcePartPath.lastIndexOf("/");
    const baseDir = slash === -1 ? "" : sourcePartPath.slice(0, slash);
    combined = baseDir === "" ? target : `${baseDir}/${target}`;
  }

  const segments: string[] = [];
  for (const segment of combined.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join("/");
}

function readRawPackageText(source: PptxSourceModel, partPath: PartPath): string | undefined {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart?.kind === "binary") return textDecoder.decode(rawPart.bytes);
  return undefined;
}

function normalizeHex(hex: string): string {
  // OOXML `srgbClr@val` / `sysClr@lastClr` は 6 桁 RRGGBB 前提。
  const normalized = hex.replace(/^#/, "").toLowerCase();
  return `#${normalized.padStart(6, "0").slice(0, 6)}`;
}

function omitColor(properties: SourceRunProperties): Omit<SourceRunProperties, "color"> {
  const withoutColor = { ...properties };
  delete withoutColor.color;
  return withoutColor;
}

function applyLuminance(hex: string, lumMod: number, lumOff: number): string {
  const { h, s, l } = hexToHsl(hex);
  return hslToHex(h, s, Math.min(1, Math.max(0, l * lumMod + lumOff)));
}

function applyTint(hex: string, tintAmount: number): string {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(
    Math.round(r + (255 - r) * tintAmount),
    Math.round(g + (255 - g) * tintAmount),
    Math.round(b + (255 - b) * tintAmount),
  );
}

function applyShade(hex: string, shadeAmount: number): string {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(
    Math.round(r * shadeAmount),
    Math.round(g * shadeAmount),
    Math.round(b * shadeAmount),
  );
}

function hexToRgb(hex: string): { readonly r: number; readonly g: number; readonly b: number } {
  const normalized = hex.replace("#", "");
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16),
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (value: number) => Math.min(255, Math.max(0, value)).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function hexToHsl(hex: string): { readonly h: number; readonly s: number; readonly l: number } {
  const { r: r255, g: g255, b: b255 } = hexToRgb(hex);
  const r = r255 / 255;
  const g = g255 / 255;
  const b = b255 / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;

  if (max === min) return { h: 0, s: 0, l };

  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h =
    max === r
      ? ((g - b) / d + (g < b ? 6 : 0)) / 6
      : max === g
        ? ((b - r) / d + 2) / 6
        : ((r - g) / d + 4) / 6;

  return { h, s, l };
}

function hslToHex(h: number, s: number, l: number): string {
  if (s === 0) {
    const value = Math.round(l * 255);
    return rgbToHex(value, value, value);
  }

  const hue2rgb = (p: number, q: number, t: number): number => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return rgbToHex(
    Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, h) * 255),
    Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
  );
}

function firstDefined<T>(...values: readonly (T | undefined)[]): T | undefined {
  return values.find((value) => value !== undefined);
}
