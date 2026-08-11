import type { Node as XmlDomNode } from "@xmldom/xmldom";
import { DOMImplementation, XMLSerializer } from "@xmldom/xmldom";
import * as emfModule from "rtf.js/dist/EMFJS.bundle.js";
import * as wmfModule from "rtf.js/dist/WMFJS.bundle.js";

import type { ImageMimeType } from "../model/tokens.js";
import type { WarningLogger } from "../warning-logger.js";

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const MM_ANISOTROPIC = 8;

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

export type MetafileConversionCache = Map<string, MetafileConversionResult>;

export function resolveMetafileImageSource(
  imageData: string,
  mimeType: ImageMimeType,
  warningLogger: WarningLogger,
  cache?: MetafileConversionCache,
): ResolvedImageSource | undefined {
  if (mimeType !== "image/emf" && mimeType !== "image/wmf") {
    return { imageData, mimeType };
  }
  const cacheKey = `${mimeType}:${imageData}`;
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
): string {
  const svg = new TextDecoder().decode(base64ToUint8Array(imageData));
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
    bytes = base64ToUint8Array(imageData);
  } catch (error: unknown) {
    return failure("invalid-data", errorMessage(error));
  }

  const validation = validateMetafile(bytes, mimeType);
  if (!validation.ok) return validation;
  const geometry = validation.geometry;

  const document = new DOMImplementation().createDocument(SVG_NAMESPACE, "svg");
  const previousDocument = Reflect.get(globalThis, "document");
  const hadDocument = Reflect.has(globalThis, "document");

  try {
    Reflect.set(globalThis, "document", document);
    const emfJs = resolveRendererApi(emfModule);
    const wmfJs = resolveRendererApi(wmfModule);
    emfJs.loggingEnabled(false);
    wmfJs.loggingEnabled(false);

    const arrayBuffer = new ArrayBuffer(bytes.byteLength);
    new Uint8Array(arrayBuffer).set(bytes);
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

    let svg = new XMLSerializer().serializeToString(unsafeMetafileNodeAssertion(svgElement));
    if (mimeType === "image/emf") {
      svg = appendEmfText(svg, bytes);
    }
    svg = normalizeMetafileSvg(svg, geometry);
    return {
      ok: true,
      imageData: uint8ArrayToBase64(new TextEncoder().encode(svg)),
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
  | { readonly ok: true; readonly geometry: MetafileGeometry }
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
    if (!SUPPORTED_EMF_RECORD_TYPES.has(type)) {
      return failure("unsupported-record", `EMF record 0x${type.toString(16)} is unsupported.`);
    }
    recordCount++;
    if (recordCount > 200_000) {
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
  return { ok: true, geometry };
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
  if ((type !== 1 && type !== 2) || headerSizeWords !== 9) {
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
  let foundEof = false;
  let maximumRecordWords = 0;
  while (offset < bytes.byteLength) {
    if (offset + 6 > bytes.byteLength) {
      return failure("invalid-data", "WMF record header is truncated.");
    }
    const sizeWords = view.getUint32(offset, true);
    const recordType = view.getUint16(offset + 4, true);
    const size = sizeWords * 2;
    if (size < 6 || offset + size > bytes.byteLength) {
      return failure(
        "invalid-data",
        `WMF record 0x${recordType.toString(16)} has an invalid size.`,
      );
    }
    if (!SUPPORTED_WMF_RECORD_TYPES.has(recordType)) {
      return failure(
        "unsupported-record",
        `WMF record 0x${recordType.toString(16)} is unsupported.`,
      );
    }
    recordCount++;
    maximumRecordWords = Math.max(maximumRecordWords, sizeWords);
    if (recordCount > 200_000) {
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

  let originX = 0;
  let originY = 0;
  let width = 0;
  let height = 0;
  let offset = headerOffset + 18;
  while (offset + 6 <= view.byteLength) {
    const size = view.getUint32(offset, true) * 2;
    const type = view.getUint16(offset + 4, true);
    if (type === 0x020b && size >= 10) {
      originY = view.getInt16(offset + 6, true);
      originX = view.getInt16(offset + 8, true);
    } else if (type === 0x020c && size >= 10) {
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

function appendEmfText(svg: string, bytes: Uint8Array): string {
  const view = dataView(bytes);
  const textElements: string[] = [];
  let textColor = "#000000";
  let offset = 0;

  while (offset + 8 <= bytes.byteLength) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.byteLength) break;

    if (type === 0x18 && size >= 12) {
      textColor = colorRefToHex(view.getUint32(offset + 8, true));
    } else if ((type === 0x53 || type === 0x54) && size >= 76) {
      const x = view.getInt32(offset + 36, true);
      const y = view.getInt32(offset + 40, true);
      const charCount = view.getUint32(offset + 44, true);
      const stringOffset = view.getUint32(offset + 48, true);
      const bytesPerCharacter = type === 0x54 ? 2 : 1;
      const start = offset + stringOffset;
      const end = start + charCount * bytesPerCharacter;
      if (start >= offset && end <= offset + size) {
        const textBytes = bytes.subarray(start, end);
        const text =
          type === 0x54
            ? new TextDecoder("utf-16le").decode(textBytes)
            : new TextDecoder("windows-1252").decode(textBytes);
        textElements.push(
          `<text x="${x}" y="${y}" fill="${textColor}" font-family="sans-serif" font-size="96">${escapeXmlText(text)}</text>`,
        );
      }
    }

    offset += size;
    if (type === 0x0e) break;
  }

  return textElements.length > 0 ? svg.replace("</svg>", `${textElements.join("")}</svg>`) : svg;
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
  if (typeof Buffer !== "undefined") {
    const bytes = Buffer.from(value, "base64");
    if (bytes.length === 0 && value.length > 0) throw new Error("Invalid base64 image data.");
    return new Uint8Array(bytes);
  }
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}
