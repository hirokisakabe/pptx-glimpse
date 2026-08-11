import type { Node as XmlDomNode } from "@xmldom/xmldom";
import { DOMImplementation, XMLSerializer } from "@xmldom/xmldom";
import * as emfModule from "rtf.js/dist/EMFJS.bundle.js";
import * as wmfModule from "rtf.js/dist/WMFJS.bundle.js";

import type { ImageMimeType } from "../model/tokens.js";
import type { WarningLogger } from "../warning-logger.js";

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const MM_ANISOTROPIC = 8;
const MAX_METAFILE_BYTES = 8 * 1024 * 1024;
const MAX_BASE64_LENGTH = Math.ceil(MAX_METAFILE_BYTES / 3) * 4;
const MAX_RECORD_BYTES = 4 * 1024 * 1024;
const MAX_BITMAP_BYTES = 2 * 1024 * 1024;
const MAX_RECORDS = 50_000;
const MAX_GEOMETRY_POINTS = 200_000;
const MAX_SVG_NODES = 100_000;
const MAX_SVG_LENGTH = 8 * 1024 * 1024;
const MAX_CACHE_ENTRIES = 16;
const MAX_CACHE_BYTES = 16 * 1024 * 1024;

interface MetafileRendererApi {
  readonly Renderer: new (data: ArrayBuffer) => {
    render(settings: Record<string, string | number>): unknown;
  };
  readonly loggingEnabled: (enabled: boolean) => void;
}

interface MetafileRendererModule {
  readonly Renderer?: MetafileRendererApi["Renderer"];
  readonly loggingEnabled?: MetafileRendererApi["loggingEnabled"];
  readonly default?: MetafileRendererApi;
  readonly EMFJS?: MetafileRendererApi;
  readonly WMFJS?: MetafileRendererApi;
}

const SUPPORTED_EMF_RECORD_TYPES = new Set([
  0x01, 0x02, 0x03, 0x05, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x11, 0x12, 0x13, 0x15, 0x16,
  0x18, 0x19, 0x1a, 0x1b, 0x21, 0x22, 0x25, 0x26, 0x27, 0x28, 0x2b, 0x2c, 0x36, 0x3a, 0x3b, 0x3c,
  0x3d, 0x3e, 0x40, 0x43, 0x44, 0x4b, 0x52, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5b, 0x5f,
]);

const SUPPORTED_WMF_RECORD_TYPES = new Set([
  0x0000, 0x001e, 0x00f7, 0x0102, 0x0106, 0x0107, 0x0127, 0x012c, 0x012d, 0x012e, 0x0142, 0x01f0,
  0x01f9, 0x0201, 0x0209, 0x020b, 0x020c, 0x020d, 0x020e, 0x020f, 0x0211, 0x0213, 0x0214, 0x0220,
  0x0234, 0x02fa, 0x02fb, 0x02fc, 0x0324, 0x0325, 0x0415, 0x0416, 0x0418, 0x041b, 0x0521, 0x0538,
  0x061c, 0x0626, 0x06ff, 0x0940, 0x0a32, 0x0b41, 0x0f43,
]);

const COUNTED_EMF_POINT_RECORD_TYPES = new Set([0x02, 0x03, 0x05, 0x55, 0x56, 0x57, 0x58, 0x59]);

export type MetafileMimeType = Extract<ImageMimeType, "image/emf" | "image/wmf">;

type MetafileConversionFailureReason = "invalid-data" | "unsupported-record" | "conversion-failed";

export type MetafileConversionResult =
  | { readonly ok: true; readonly imageData: string; readonly mimeType: "image/svg+xml" }
  | {
      readonly ok: false;
      readonly reason: MetafileConversionFailureReason;
      readonly message: string;
    };

export interface ResolvedImageSource {
  readonly imageData: string;
  readonly mimeType: ImageMimeType;
}

export interface MetafileConversionCache {
  readonly size: number;
  get(key: string): MetafileConversionResult | undefined;
  set(key: string, result: MetafileConversionResult): void;
}

export class BoundedMetafileConversionCache implements MetafileConversionCache {
  readonly #entries = new Map<
    string,
    { readonly result: MetafileConversionResult; readonly bytes: number }
  >();
  #bytes = 0;

  get size(): number {
    return this.#entries.size;
  }

  get(key: string): MetafileConversionResult | undefined {
    const entry = this.#entries.get(key);
    if (entry === undefined) return undefined;
    this.#entries.delete(key);
    this.#entries.set(key, entry);
    return entry.result;
  }

  set(key: string, result: MetafileConversionResult): void {
    const bytes = conversionResultBytes(result) + key.length * 2;
    if (bytes > MAX_CACHE_BYTES) return;
    const previous = this.#entries.get(key);
    if (previous !== undefined) {
      this.#bytes -= previous.bytes;
      this.#entries.delete(key);
    }
    this.#entries.set(key, { result, bytes });
    this.#bytes += bytes;
    while (this.#entries.size > MAX_CACHE_ENTRIES || this.#bytes > MAX_CACHE_BYTES) {
      const oldestKey = this.#entries.keys().next().value;
      if (oldestKey === undefined) break;
      const oldest = this.#entries.get(oldestKey);
      this.#entries.delete(oldestKey);
      this.#bytes -= oldest?.bytes ?? 0;
    }
  }
}

export function resolveMetafileImageSource(
  imageData: string,
  mimeType: ImageMimeType,
  warningLogger: WarningLogger,
  cache?: MetafileConversionCache,
): ResolvedImageSource | undefined {
  if (mimeType !== "image/emf" && mimeType !== "image/wmf") {
    return { imageData, mimeType };
  }
  const cacheKey = metafileCacheKey(imageData, mimeType);
  let result = cache?.get(cacheKey);
  if (result === undefined) {
    result = convertMetafileToSvgData(imageData, mimeType);
    cache?.set(cacheKey, result);
  }
  if (result.ok) return result;

  warningLogger.warn(
    "image.metafile-conversion",
    `${mimeType === "image/emf" ? "EMF" : "WMF"} conversion ${result.reason}: ${result.message}`,
  );
  return undefined;
}

export function inlineSvgData(
  imageData: string,
  attributes: Readonly<Record<string, string | number>>,
  idNamespace = "",
): string {
  const svg = namespaceSvgIds(new TextDecoder().decode(base64ToUint8Array(imageData)), idNamespace);
  const openingEnd = svg.indexOf(">");
  if (!svg.startsWith("<svg") || openingEnd < 0) {
    throw new Error("Converted metafile SVG root is invalid.");
  }
  const opening = svg
    .slice(0, openingEnd)
    .replace(/\s(?:x|y|width|height|preserveAspectRatio)="[^"]*"/g, "");
  const renderedAttributes = Object.entries(attributes)
    .map(([name, value]) => {
      if (!INLINE_SVG_ATTRIBUTES.has(name)) {
        throw new Error(`Unsupported inline SVG attribute '${name}'.`);
      }
      return ` ${name}="${escapeXmlAttribute(String(value))}"`;
    })
    .join("");
  return `${opening}${renderedAttributes}>${svg.slice(openingEnd + 1)}`;
}

export function convertMetafileToSvgData(
  imageData: string,
  mimeType: MetafileMimeType,
): MetafileConversionResult {
  let bytes: Uint8Array;
  try {
    if (imageData.length > MAX_BASE64_LENGTH) {
      return failure("invalid-data", "Metafile encoded input limit exceeded.");
    }
    bytes = base64ToUint8Array(imageData);
  } catch (error: unknown) {
    return failure("invalid-data", errorMessage(error));
  }

  if (bytes.byteLength > MAX_METAFILE_BYTES) {
    return failure("invalid-data", "Metafile decoded byte limit exceeded.");
  }

  const validation = validateMetafile(bytes, mimeType);
  if (!validation.ok) return validation;
  const geometry = validation.geometry;
  let renderBytes = bytes;
  if (mimeType === "image/emf") {
    const normalizedRecords = normalizeEmfRestoreDcRecords(bytes);
    if (!normalizedRecords.ok) return normalizedRecords;
    renderBytes = normalizedRecords.bytes;
  }

  const document = new DOMImplementation().createDocument(SVG_NAMESPACE, "svg");
  const previousDocument = Reflect.get(globalThis, "document");
  const hadDocument = Reflect.has(globalThis, "document");

  try {
    Reflect.set(globalThis, "document", document);
    const emfJs = resolveRendererApi(emfModule);
    const wmfJs = resolveRendererApi(wmfModule);
    emfJs.loggingEnabled(false);
    wmfJs.loggingEnabled(false);

    const arrayBuffer = new ArrayBuffer(renderBytes.byteLength);
    new Uint8Array(arrayBuffer).set(renderBytes);
    const svgElement =
      mimeType === "image/emf"
        ? new emfJs.Renderer(arrayBuffer).render({
            width: String(geometry.width),
            height: String(geometry.height),
            wExt: geometry.width,
            hExt: geometry.height,
            xExt: geometry.width,
            yExt: geometry.height,
            mapMode: MM_ANISOTROPIC,
          })
        : new wmfJs.Renderer(arrayBuffer).render({
            width: String(geometry.width),
            height: String(geometry.height),
            xExt: geometry.width,
            yExt: geometry.height,
            mapMode: MM_ANISOTROPIC,
          });

    const svgNode = unsafeMetafileNodeAssertion(svgElement);
    if (countSvgNodes(svgNode) > MAX_SVG_NODES) {
      return failure("conversion-failed", "Converted metafile SVG node limit exceeded.");
    }
    let svg = new XMLSerializer().serializeToString(svgNode);
    if (validation.emfTextSvg !== undefined) svg = appendSvgContent(svg, validation.emfTextSvg);
    svg = normalizeMetafileSvg(svg, geometry);
    svg = sanitizeXml10(svg);
    if (svg.length > MAX_SVG_LENGTH) {
      return failure("conversion-failed", "Converted metafile SVG output limit exceeded.");
    }
    const svgBytes = new TextEncoder().encode(svg);
    if (svgBytes.byteLength > MAX_SVG_LENGTH) {
      return failure("conversion-failed", "Converted metafile SVG output limit exceeded.");
    }
    return {
      ok: true,
      imageData: uint8ArrayToBase64(svgBytes),
      mimeType: "image/svg+xml",
    };
  } catch (error: unknown) {
    return failure("conversion-failed", errorMessage(error));
  } finally {
    if (hadDocument) {
      Reflect.set(globalThis, "document", previousDocument);
    } else {
      Reflect.deleteProperty(globalThis, "document");
    }
  }
}

interface MetafileGeometry {
  readonly originX: number;
  readonly originY: number;
  readonly width: number;
  readonly height: number;
}

type MetafileValidationResult =
  | { readonly ok: true; readonly geometry: MetafileGeometry; readonly emfTextSvg?: string }
  | Exclude<MetafileConversionResult, { readonly ok: true }>;

const INLINE_SVG_ATTRIBUTES = new Set(["x", "y", "width", "height", "preserveAspectRatio"]);

function normalizeMetafileSvg(svg: string, geometry: MetafileGeometry): string {
  return svg
    .replace(
      / viewBox="[^"]*"/,
      ` viewBox="${geometry.originX} ${geometry.originY} ${geometry.width} ${geometry.height}"`,
    )
    .replace(/ filter="url\(#undefined\)"/g, "")
    .replace(/font-family="(?:sans-serif|Helvetica)"/g, 'font-family="Noto Sans"');
}

function unsafeMetafileNodeAssertion(value: unknown): XmlDomNode {
  // The bundled rtf.js declarations expose a browser SVGElement even when an xmldom Document
  // supplies the runtime node. This assertion is isolated at that external-library boundary.
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion
  return value as XmlDomNode;
}

function resolveRendererApi(module: MetafileRendererModule): MetafileRendererApi {
  if (module.Renderer !== undefined && module.loggingEnabled !== undefined) {
    return { Renderer: module.Renderer, loggingEnabled: module.loggingEnabled };
  }
  const api = module.default ?? module.EMFJS ?? module.WMFJS;
  if (api === undefined) throw new Error("Metafile renderer module has no compatible API.");
  return api;
}

function validateMetafile(bytes: Uint8Array, mimeType: MetafileMimeType): MetafileValidationResult {
  return mimeType === "image/emf" ? validateEmf(bytes) : validateWmf(bytes);
}

function validateEmf(bytes: Uint8Array): MetafileValidationResult {
  if (bytes.byteLength < 88) return failure("invalid-data", "EMF header is truncated.");
  const view = dataView(bytes);
  if (view.getUint32(0, true) !== 1 || view.getUint32(40, true) !== 0x464d4520) {
    return failure("invalid-data", "EMF header signature is invalid.");
  }
  const headerSize = view.getUint32(4, true);
  const declaredBytes = view.getUint32(48, true);
  const declaredRecords = view.getUint32(52, true);
  if (headerSize < 88 || headerSize % 4 !== 0 || headerSize > bytes.byteLength) {
    return failure("invalid-data", "EMF header size is invalid.");
  }
  if (declaredBytes !== bytes.byteLength) {
    return failure("invalid-data", "EMF declared byte size does not match the stream.");
  }

  const geometry = emfHeaderGeometry(view);
  if (geometry === undefined) return failure("invalid-data", "EMF header bounds are empty.");

  let offset = 0;
  let recordCount = 0;
  let geometryPointCount = 0;
  let foundEof = false;
  while (offset < bytes.byteLength) {
    if (offset + 8 > bytes.byteLength) {
      return failure("invalid-data", "EMF record header is truncated.");
    }
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || size % 4 !== 0 || offset + size > bytes.byteLength) {
      return failure("invalid-data", `EMF record 0x${type.toString(16)} has an invalid size.`);
    }
    if (size > MAX_RECORD_BYTES) {
      return failure("invalid-data", "EMF record payload limit exceeded.");
    }
    if (!SUPPORTED_EMF_RECORD_TYPES.has(type)) {
      return failure("unsupported-record", `EMF record 0x${type.toString(16)} is unsupported.`);
    }
    if (type === 0x0e) {
      const eofValidation = validateEmfEof(view, offset, size);
      if (eofValidation !== undefined) return eofValidation;
    }
    const allocationValidation = validateEmfRecordAllocations(view, offset, type, size);
    if (!allocationValidation.ok) return allocationValidation;
    geometryPointCount += allocationValidation.geometryPoints;
    if (geometryPointCount > MAX_GEOMETRY_POINTS) {
      return failure("invalid-data", "EMF geometry point limit exceeded.");
    }
    recordCount++;
    if (recordCount > MAX_RECORDS) {
      return failure("invalid-data", "EMF record limit exceeded.");
    }
    offset += size;
    if (type === 0x0e) {
      foundEof = true;
      break;
    }
  }
  if (!foundEof) return failure("invalid-data", "EMF EOF record is missing.");
  if (offset !== bytes.byteLength) {
    return failure("invalid-data", "EMF stream has trailing bytes after EOF.");
  }
  if (recordCount !== declaredRecords) {
    return failure("invalid-data", "EMF declared record count does not match the stream.");
  }
  const text = extractEmfText(bytes, geometry);
  if (!text.ok) return text;
  return {
    ok: true,
    geometry,
    ...(text.svg.length > 0 ? { emfTextSvg: text.svg } : {}),
  };
}

function validateWmf(bytes: Uint8Array): MetafileValidationResult {
  if (bytes.byteLength < 18) return failure("invalid-data", "WMF header is truncated.");
  const view = dataView(bytes);
  const placeable = view.getUint32(0, true) === 0x9ac6cdd7;
  const headerOffset = placeable ? 22 : 0;
  if (bytes.byteLength < headerOffset + 18) {
    return failure("invalid-data", "WMF meta header is truncated.");
  }
  const type = view.getUint16(headerOffset, true);
  const headerSizeWords = view.getUint16(headerOffset + 2, true);
  const version = view.getUint16(headerOffset + 4, true);
  const parameterCount = view.getUint16(headerOffset + 16, true);
  if (
    (type !== 1 && type !== 2) ||
    headerSizeWords !== 9 ||
    (version !== 0x0100 && version !== 0x0300) ||
    parameterCount !== 0
  ) {
    return failure("invalid-data", "WMF meta header is invalid.");
  }
  const declaredSizeBytes = view.getUint32(headerOffset + 6, true) * 2;
  if (declaredSizeBytes !== bytes.byteLength - headerOffset) {
    return failure("invalid-data", "WMF declared byte size does not match the stream.");
  }
  if (placeable && !hasValidPlaceableChecksum(view)) {
    return failure("invalid-data", "WMF placeable header checksum is invalid.");
  }

  let offset = headerOffset + headerSizeWords * 2;
  let recordCount = 0;
  let geometryPointCount = 0;
  let foundEof = false;
  let maximumRecordWords = 0;
  while (offset < bytes.byteLength) {
    if (offset + 6 > bytes.byteLength) {
      return failure("invalid-data", "WMF record header is truncated.");
    }
    const sizeWords = view.getUint32(offset, true);
    const recordType = view.getUint16(offset + 4, true);
    const size = sizeWords * 2;
    if (recordType === 0 && sizeWords !== 3) {
      return failure("invalid-data", "WMF EOF record has an invalid declared size.");
    }
    if (size < 6 || offset + size > bytes.byteLength) {
      return failure(
        "invalid-data",
        `WMF record 0x${recordType.toString(16)} has an invalid size.`,
      );
    }
    if (size > MAX_RECORD_BYTES) {
      return failure("invalid-data", "WMF record payload limit exceeded.");
    }
    const allocationValidation = validateWmfRecordAllocations(view, offset, recordType, size);
    if (!allocationValidation.ok) return allocationValidation;
    geometryPointCount += allocationValidation.geometryPoints;
    if (geometryPointCount > MAX_GEOMETRY_POINTS) {
      return failure("invalid-data", "WMF geometry point limit exceeded.");
    }
    if (!SUPPORTED_WMF_RECORD_TYPES.has(recordType)) {
      return failure(
        "unsupported-record",
        `WMF record 0x${recordType.toString(16)} is unsupported.`,
      );
    }
    recordCount++;
    maximumRecordWords = Math.max(maximumRecordWords, sizeWords);
    if (recordCount > MAX_RECORDS) {
      return failure("invalid-data", "WMF record limit exceeded.");
    }
    offset += size;
    if (recordType === 0) {
      foundEof = true;
      break;
    }
  }
  if (!foundEof) return failure("invalid-data", "WMF EOF record is missing.");
  if (offset !== bytes.byteLength) {
    return failure("invalid-data", "WMF stream has trailing bytes after EOF.");
  }
  if (view.getUint32(headerOffset + 12, true) !== maximumRecordWords) {
    return failure("invalid-data", "WMF declared maximum record size does not match the stream.");
  }
  const geometry = wmfGeometry(view, headerOffset, placeable);
  if (geometry === undefined) {
    return failure("invalid-data", "WMF has no usable placeable bounds or window extent.");
  }
  return { ok: true, geometry };
}

function emfHeaderGeometry(view: DataView): MetafileGeometry | undefined {
  const bounds = rectangleGeometry(view, 8);
  if (bounds !== undefined) return bounds;
  return rectangleGeometry(view, 24);
}

function rectangleGeometry(view: DataView, offset: number): MetafileGeometry | undefined {
  const originX = view.getInt32(offset, true);
  const originY = view.getInt32(offset + 4, true);
  const width = view.getInt32(offset + 8, true) - originX;
  const height = view.getInt32(offset + 12, true) - originY;
  return width > 0 && height > 0 ? { originX, originY, width, height } : undefined;
}

function hasValidPlaceableChecksum(view: DataView): boolean {
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
  return checksum === view.getUint16(20, true);
}

function wmfGeometry(
  view: DataView,
  headerOffset: number,
  placeable: boolean,
): MetafileGeometry | undefined {
  if (placeable) {
    const originX = view.getInt16(6, true);
    const originY = view.getInt16(8, true);
    const width = view.getInt16(10, true) - originX;
    const height = view.getInt16(12, true) - originY;
    if (width > 0 && height > 0) return { originX, originY, width, height };
  }

  let width = 0;
  let height = 0;
  let offset = headerOffset + 18;
  while (offset + 6 <= view.byteLength) {
    const size = view.getUint32(offset, true) * 2;
    const type = view.getUint16(offset + 4, true);
    if (type === 0x020c && size >= 10) {
      height = Math.abs(view.getInt16(offset + 6, true));
      width = Math.abs(view.getInt16(offset + 8, true));
      break;
    }
    offset += size;
    if (type === 0) break;
  }
  // WMFJS applies META_SETWINDOWORG while rendering, so its output device coordinates begin at
  // zero even when the logical stream window has a non-zero origin.
  return width > 0 && height > 0 ? { originX: 0, originY: 0, width, height } : undefined;
}

interface EmfFontState {
  readonly height: number;
  readonly width: number;
  readonly escapement: number;
  readonly orientation: number;
  readonly weight: number;
  readonly italic: boolean;
  readonly faceName: string;
}

interface EmfMappingState {
  mapMode: number;
  windowX: number;
  windowY: number;
  windowWidth: number;
  windowHeight: number;
  viewportX: number;
  viewportY: number;
  viewportWidth: number;
  viewportHeight: number;
}

interface EmfTextState {
  readonly mapping: EmfMappingState;
  readonly selectedFont: EmfFontState | undefined;
  readonly textColor: string;
  readonly textAlign: number;
}

type EmfTextExtractionResult =
  | { readonly ok: true; readonly svg: string }
  | Exclude<MetafileConversionResult, { readonly ok: true }>;

function extractEmfText(bytes: Uint8Array, geometry: MetafileGeometry): EmfTextExtractionResult {
  const view = dataView(bytes);
  const textElements: string[] = [];
  const fonts = new Map<number, EmfFontState>();
  let state: EmfTextState = {
    selectedFont: undefined,
    textColor: "#000000",
    textAlign: 0,
    mapping: {
      mapMode: MM_ANISOTROPIC,
      windowX: 0,
      windowY: 0,
      windowWidth: geometry.width,
      windowHeight: geometry.height,
      viewportX: 0,
      viewportY: 0,
      viewportWidth: geometry.width,
      viewportHeight: geometry.height,
    },
  };
  const savedStates: EmfTextState[] = [];
  let offset = 0;

  while (offset + 8 <= bytes.byteLength) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.byteLength) break;

    if (type === 0x09 && size >= 16) {
      state.mapping.windowWidth = view.getInt32(offset + 8, true);
      state.mapping.windowHeight = view.getInt32(offset + 12, true);
    } else if (type === 0x0a && size >= 16) {
      state.mapping.windowX = view.getInt32(offset + 8, true);
      state.mapping.windowY = view.getInt32(offset + 12, true);
    } else if (type === 0x0b && size >= 16) {
      state.mapping.viewportWidth = view.getInt32(offset + 8, true);
      state.mapping.viewportHeight = view.getInt32(offset + 12, true);
    } else if (type === 0x0c && size >= 16) {
      state.mapping.viewportX = view.getInt32(offset + 8, true);
      state.mapping.viewportY = view.getInt32(offset + 12, true);
    } else if (type === 0x11 && size >= 12) {
      state.mapping.mapMode = view.getInt32(offset + 8, true);
    } else if (type === 0x16 && size >= 12) {
      state = { ...state, textAlign: view.getUint32(offset + 8, true) };
    } else if (type === 0x18 && size >= 12) {
      state = { ...state, textColor: colorRefToHex(view.getUint32(offset + 8, true)) };
    } else if (type === 0x21) {
      savedStates.push(cloneEmfTextState(state));
    } else if (type === 0x22 && size >= 12) {
      const restored = restoreEmfTextState(savedStates, view.getInt32(offset + 8, true));
      if (restored === undefined) {
        return failure("invalid-data", "EMF RestoreDC references an invalid saved state.");
      }
      state = restored;
    } else if (type === 0x52) {
      const font = readEmfFont(view, offset, size);
      if (!font.ok) return font;
      fonts.set(view.getUint32(offset + 8, true), font.font);
    } else if (type === 0x25 && size >= 12) {
      const index = view.getUint32(offset + 8, true);
      if (fonts.has(index)) state = { ...state, selectedFont: fonts.get(index) };
    } else if (type === 0x28 && size >= 12) {
      fonts.delete(view.getUint32(offset + 8, true));
    } else if ((type === 0x53 || type === 0x54) && size >= 76) {
      const text = renderEmfTextRecord(bytes, view, offset, size, type, state);
      if (!text.ok) return text;
      textElements.push(text.svg);
    }

    offset += size;
    if (type === 0x0e) break;
  }

  return { ok: true, svg: textElements.join("") };
}

function readEmfFont(
  view: DataView,
  offset: number,
  size: number,
):
  | { readonly ok: true; readonly font: EmfFontState }
  | Exclude<MetafileConversionResult, { readonly ok: true }> {
  if (size < 104) return failure("invalid-data", "EMF font record is truncated.");
  const width = view.getInt32(offset + 16, true);
  const orientation = view.getInt32(offset + 24, true);
  const escapement = view.getInt32(offset + 20, true);
  if (width !== 0 || (orientation !== 0 && orientation !== escapement)) {
    return failure("unsupported-record", "EMF text font width/orientation is unsupported.");
  }
  const faceBytes = new Uint8Array(view.buffer, view.byteOffset + offset + 40, 64);
  const faceName =
    sanitizeXml10(new TextDecoder("utf-16le").decode(faceBytes).split("\0", 1)[0]) || "Noto Sans";
  return {
    ok: true,
    font: {
      height: view.getInt32(offset + 12, true),
      width,
      escapement,
      orientation,
      weight: view.getInt32(offset + 28, true),
      italic: view.getUint8(offset + 32) !== 0,
      faceName,
    },
  };
}

function renderEmfTextRecord(
  bytes: Uint8Array,
  view: DataView,
  offset: number,
  size: number,
  type: number,
  state: {
    readonly mapping: EmfMappingState;
    readonly textAlign: number;
    readonly textColor: string;
    readonly selectedFont: EmfFontState | undefined;
  },
): EmfTextExtractionResult {
  const graphicsMode = view.getUint32(offset + 24, true);
  const scaleX = view.getFloat32(offset + 28, true);
  const scaleY = view.getFloat32(offset + 32, true);
  const options = view.getUint32(offset + 52, true);
  if (
    state.mapping.mapMode !== MM_ANISOTROPIC ||
    graphicsMode !== 1 ||
    scaleX !== 1 ||
    scaleY !== 1 ||
    state.textAlign !== 0 ||
    options !== 0 ||
    state.mapping.windowWidth === 0 ||
    state.mapping.windowHeight === 0
  ) {
    return failure("unsupported-record", "EMF text mapping/alignment/options are unsupported.");
  }

  const charCount = view.getUint32(offset + 44, true);
  const stringOffset = view.getUint32(offset + 48, true);
  const bytesPerCharacter = type === 0x54 ? 2 : 1;
  const start = offset + stringOffset;
  const end = start + charCount * bytesPerCharacter;
  if (start < offset || end > offset + size) {
    return failure("invalid-data", "EMF text string range is invalid.");
  }
  const textBytes = bytes.subarray(start, end);
  const text = sanitizeXml10(
    type === 0x54
      ? new TextDecoder("utf-16le").decode(textBytes)
      : new TextDecoder("windows-1252").decode(textBytes),
  );
  const font = state.selectedFont ?? {
    height: -80,
    width: 0,
    escapement: 0,
    orientation: 0,
    weight: 400,
    italic: false,
    faceName: "Noto Sans",
  };
  const x = mapEmfCoordinate(
    view.getInt32(offset + 36, true),
    state.mapping.windowX,
    state.mapping.windowWidth,
    state.mapping.viewportX,
    state.mapping.viewportWidth,
  );
  const y = mapEmfCoordinate(
    view.getInt32(offset + 40, true),
    state.mapping.windowY,
    state.mapping.windowHeight,
    state.mapping.viewportY,
    state.mapping.viewportHeight,
  );
  const fontSize = Math.abs(
    mapEmfLength(font.height, state.mapping.windowHeight, state.mapping.viewportHeight),
  );
  if (!Number.isFinite(fontSize) || fontSize === 0) {
    return failure("unsupported-record", "EMF text font height cannot be mapped.");
  }

  const offDx = view.getUint32(offset + 72, true);
  let dxAttribute = "";
  if (offDx !== 0) {
    const dxStart = offset + offDx;
    const dxEnd = dxStart + charCount * 4;
    if (dxStart < offset || dxEnd > offset + size) {
      return failure("invalid-data", "EMF text advance range is invalid.");
    }
    const advances = Array.from({ length: Math.max(0, charCount - 1) }, (_, index) =>
      mapEmfLength(
        view.getInt32(dxStart + index * 4, true),
        state.mapping.windowWidth,
        state.mapping.viewportWidth,
      ),
    );
    dxAttribute = ` dx="${[0, ...advances].join(" ")}"`;
  }

  const rotation =
    font.escapement === 0 ? "" : ` transform="rotate(${-font.escapement / 10} ${x} ${y})"`;
  const weight = font.weight >= 700 ? ' font-weight="bold"' : "";
  const italic = font.italic ? ' font-style="italic"' : "";
  return {
    ok: true,
    svg: `<text x="${x}" y="${y}"${dxAttribute} fill="${state.textColor}" font-family="${escapeXmlAttribute(font.faceName)}" font-size="${fontSize}" dominant-baseline="hanging"${weight}${italic}${rotation}>${escapeXmlText(text)}</text>`,
  };
}

function mapEmfCoordinate(
  value: number,
  windowOrigin: number,
  windowExtent: number,
  viewportOrigin: number,
  viewportExtent: number,
): number {
  return ((value - windowOrigin) * viewportExtent) / windowExtent + viewportOrigin;
}

function mapEmfLength(value: number, windowExtent: number, viewportExtent: number): number {
  return (value * viewportExtent) / windowExtent;
}

function appendSvgContent(svg: string, content: string): string {
  return svg.replace("</svg>", `${content}</svg>`);
}

function cloneEmfTextState(state: EmfTextState): EmfTextState {
  return { ...state, mapping: { ...state.mapping } };
}

function restoreEmfTextState(
  savedStates: EmfTextState[],
  relativeOrAbsolute: number,
): EmfTextState | undefined {
  if (relativeOrAbsolute === 0) return undefined;
  const index =
    relativeOrAbsolute < 0 ? savedStates.length + relativeOrAbsolute : relativeOrAbsolute - 1;
  const restored = savedStates[index];
  if (restored === undefined) return undefined;
  savedStates.length = index;
  return cloneEmfTextState(restored);
}

function normalizeEmfRestoreDcRecords(
  bytes: Uint8Array,
):
  | { readonly ok: true; readonly bytes: Uint8Array }
  | Exclude<MetafileConversionResult, { readonly ok: true }> {
  const view = dataView(bytes);
  if (!emfNeedsRestoreNormalization(view)) return { ok: true, bytes };
  const records: Uint8Array[] = [];
  let savedDepth = 0;
  let normalizedRecordCount = 0;
  let normalizedBytes = 0;
  let offset = 0;
  while (offset < bytes.byteLength) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (type === 0x21) savedDepth++;
    if (type !== 0x22) {
      const record = bytes.subarray(offset, offset + size);
      records.push(record);
      normalizedBytes += record.byteLength;
      normalizedRecordCount++;
    } else {
      const restoreValue = view.getInt32(offset + 8, true);
      const restoreCount = restoreValue < 0 ? -restoreValue : savedDepth - restoreValue + 1;
      for (let index = 0; index < restoreCount; index++) {
        const record = new Uint8Array(12);
        const recordView = dataView(record);
        recordView.setUint32(0, 0x22, true);
        recordView.setUint32(4, 12, true);
        recordView.setInt32(8, -1, true);
        records.push(record);
        normalizedBytes += record.byteLength;
        normalizedRecordCount++;
      }
      savedDepth -= restoreCount;
    }
    if (normalizedBytes > MAX_METAFILE_BYTES || normalizedRecordCount > MAX_RECORDS) {
      return failure("invalid-data", "Normalized EMF RestoreDC limit exceeded.");
    }
    offset += size;
  }
  if (records.length === 0) return { ok: true, bytes };
  const normalized = new Uint8Array(normalizedBytes);
  let normalizedOffset = 0;
  for (const record of records) {
    normalized.set(record, normalizedOffset);
    normalizedOffset += record.byteLength;
  }
  const normalizedView = dataView(normalized);
  normalizedView.setUint32(48, normalized.byteLength, true);
  normalizedView.setUint32(52, normalizedRecordCount, true);
  return { ok: true, bytes: normalized };
}

function emfNeedsRestoreNormalization(view: DataView): boolean {
  let offset = 0;
  while (offset < view.byteLength) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (type === 0x22 && view.getInt32(offset + 8, true) !== -1) return true;
    offset += size;
  }
  return false;
}

function validateEmfEof(
  view: DataView,
  offset: number,
  size: number,
): Exclude<MetafileConversionResult, { readonly ok: true }> | undefined {
  if (size < 20 || view.getUint32(offset + 16, true) !== size) {
    return failure("invalid-data", "EMF EOF record has an invalid declared size.");
  }
  const paletteEntries = view.getUint32(offset + 8, true);
  const paletteOffset = view.getUint32(offset + 12, true);
  if (paletteEntries === 0) return undefined;
  const paletteEnd = paletteOffset + paletteEntries * 4;
  if (
    paletteEntries > MAX_GEOMETRY_POINTS ||
    paletteOffset < 20 ||
    paletteOffset % 4 !== 0 ||
    !Number.isSafeInteger(paletteEnd) ||
    paletteEnd > size
  ) {
    return failure("invalid-data", "EMF EOF palette declaration is invalid.");
  }
  return undefined;
}

function validateEmfRecordAllocations(
  view: DataView,
  offset: number,
  type: number,
  size: number,
): RecordAllocationValidation {
  let geometryPoints = 0;
  const countOffset = COUNTED_EMF_POINT_RECORD_TYPES.has(type) ? 24 : undefined;
  if (countOffset !== undefined) {
    if (size < countOffset + 4) {
      return failure("invalid-data", "EMF geometry point declaration is invalid.");
    }
    const count = view.getUint32(offset + countOffset, true);
    const pointBytes = type >= 0x55 ? 4 : 8;
    if (count > MAX_GEOMETRY_POINTS) {
      return failure("invalid-data", "EMF geometry point limit exceeded.");
    }
    if (28 + count * pointBytes > size) {
      return failure("invalid-data", "EMF geometry point declaration is invalid.");
    }
    geometryPoints += count;
  }
  if (type === 0x08 || type === 0x5b) {
    if (size < 32) {
      return failure("invalid-data", "EMF geometry point declaration is invalid.");
    }
    const polygonCount = view.getUint32(offset + 24, true);
    const pointCount = view.getUint32(offset + 28, true);
    const pointBytes = type === 0x5b ? 4 : 8;
    if (
      polygonCount > MAX_GEOMETRY_POINTS ||
      pointCount > MAX_GEOMETRY_POINTS ||
      32 + polygonCount * 4 + pointCount * pointBytes > size
    ) {
      return failure("invalid-data", "EMF geometry point declaration is invalid.");
    }
    let declaredPointCount = 0;
    for (let index = 0; index < polygonCount; index++) {
      declaredPointCount += view.getUint32(offset + 32 + index * 4, true);
      if (declaredPointCount > MAX_GEOMETRY_POINTS) {
        return failure("invalid-data", "EMF geometry point limit exceeded.");
      }
    }
    if (declaredPointCount !== pointCount) {
      return failure("invalid-data", "EMF geometry point declaration is invalid.");
    }
    geometryPoints += pointCount;
  }
  if (type === 0x22 && size < 12) {
    return failure("invalid-data", "EMF RestoreDC record is truncated.");
  }
  if (type === 0x4b && size > MAX_BITMAP_BYTES) {
    return failure("invalid-data", "EMF bitmap payload limit exceeded.");
  }
  return { ok: true, geometryPoints };
}

function validateWmfRecordAllocations(
  view: DataView,
  offset: number,
  type: number,
  size: number,
): RecordAllocationValidation {
  let geometryPoints = 0;
  if ((type === 0x0324 || type === 0x0325) && size >= 8) {
    geometryPoints = view.getUint16(offset + 6, true);
    if (geometryPoints > MAX_GEOMETRY_POINTS || 8 + geometryPoints * 4 > size) {
      return failure("invalid-data", "WMF geometry point limit exceeded.");
    }
  }
  if (type === 0x0538 && size >= 8) {
    const polygonCount = view.getUint16(offset + 6, true);
    if (polygonCount > MAX_GEOMETRY_POINTS || 8 + polygonCount * 2 > size) {
      return failure("invalid-data", "WMF geometry point declaration is invalid.");
    }
    let pointCount = 0;
    for (let index = 0; index < polygonCount; index++) {
      pointCount += view.getUint16(offset + 8 + index * 2, true);
      if (pointCount > MAX_GEOMETRY_POINTS) {
        return failure("invalid-data", "WMF geometry point limit exceeded.");
      }
    }
    if (8 + polygonCount * 2 + pointCount * 4 > size) {
      return failure("invalid-data", "WMF geometry point declaration is invalid.");
    }
    geometryPoints = pointCount;
  }
  if ((type === 0x0940 || type === 0x0b41 || type === 0x0f43) && size > MAX_BITMAP_BYTES) {
    return failure("invalid-data", "WMF bitmap payload limit exceeded.");
  }
  return { ok: true, geometryPoints };
}

type RecordAllocationValidation =
  | { readonly ok: true; readonly geometryPoints: number }
  | Exclude<MetafileConversionResult, { readonly ok: true }>;

function countSvgNodes(root: XmlDomNode): number {
  let count = 0;
  const pending: XmlDomNode[] = [root];
  while (pending.length > 0) {
    const node = pending.pop();
    if (node === undefined) break;
    count++;
    if (count > MAX_SVG_NODES) return count;
    for (let child = node.firstChild; child !== null; child = child.nextSibling)
      pending.push(child);
  }
  return count;
}

function sanitizeXml10(value: string): string {
  let sanitized = "";
  for (const character of value) {
    const codePoint = character.codePointAt(0) ?? 0;
    sanitized +=
      codePoint === 0x09 ||
      codePoint === 0x0a ||
      codePoint === 0x0d ||
      (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
      (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
      (codePoint >= 0x10000 && codePoint <= 0x10ffff)
        ? character
        : "\uFFFD";
  }
  return sanitized;
}

function namespaceSvgIds(svg: string, namespace: string): string {
  if (namespace.length === 0) return svg;
  const ids = new Map<string, string>();
  const withIds = svg.replace(/\bid="([^"]+)"/g, (_match, id: string) => {
    const namespaced = `${namespace}${id}`;
    ids.set(id, namespaced);
    return `id="${namespaced}"`;
  });
  if (ids.size === 0) return withIds;
  return withIds
    .replace(/url\(#([^)]+)\)/g, (match, id: string) =>
      ids.has(id) ? `url(#${ids.get(id)})` : match,
    )
    .replace(/\b(xlink:href|href)="#([^"]+)"/g, (match, name: string, id: string) =>
      ids.has(id) ? `${name}="#${ids.get(id)}"` : match,
    );
}

function metafileCacheKey(imageData: string, mimeType: MetafileMimeType): string {
  let first = 0x811c9dc5;
  let second = 0x9e3779b9;
  for (let index = 0; index < imageData.length; index++) {
    const code = imageData.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193);
    second = Math.imul(second ^ code, 0x85ebca6b);
  }
  return `${mimeType}:${imageData.length}:${first >>> 0}:${second >>> 0}:${imageData.slice(0, 64)}:${imageData.slice(-64)}`;
}

function conversionResultBytes(result: MetafileConversionResult): number {
  return result.ok ? result.imageData.length : result.message.length + result.reason.length;
}

function colorRefToHex(value: number): string {
  const red = value & 0xff;
  const green = (value >>> 8) & 0xff;
  const blue = (value >>> 16) & 0xff;
  return `#${hexByte(red)}${hexByte(green)}${hexByte(blue)}`;
}

function hexByte(value: number): string {
  return value.toString(16).padStart(2, "0").toUpperCase();
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeXmlAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function dataView(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function failure(
  reason: MetafileConversionFailureReason,
  message: string,
): Exclude<MetafileConversionResult, { readonly ok: true }> {
  return { ok: false, reason, message };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function base64ToUint8Array(value: string): Uint8Array {
  validateBase64(value);
  if (typeof Buffer !== "undefined") {
    const bytes = Buffer.from(value, "base64");
    if (bytes.length === 0 && value.length > 0) throw new Error("Invalid base64 image data.");
    return new Uint8Array(bytes);
  }
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function validateBase64(value: string): void {
  if (value.length % 4 !== 0) throw new Error("Invalid base64 image data.");
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0;
  for (let index = 0; index < value.length - padding; index++) {
    const code = value.charCodeAt(index);
    const valid =
      (code >= 0x41 && code <= 0x5a) ||
      (code >= 0x61 && code <= 0x7a) ||
      (code >= 0x30 && code <= 0x39) ||
      code === 0x2b ||
      code === 0x2f;
    if (!valid) throw new Error("Invalid base64 image data.");
  }
  for (let index = value.length - padding; index < value.length; index++) {
    if (value.charCodeAt(index) !== 0x3d) throw new Error("Invalid base64 image data.");
  }
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}
