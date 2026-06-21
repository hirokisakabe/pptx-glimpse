/**
 * `writePptx(source)` — CleanDoc source writer の最初の no-edit slice。
 *
 * この writer は structural round-trip を目的に、reader が保持した raw package
 * material / media bytes / package bookkeeping を PPTX ZIP として再構成する。
 * edited writer behavior や node-level XML splicing は後続 slice の責務。
 */

import { zipSync } from "fflate";

import type {
  CleanDocSource,
  PartPath,
  PartRelationships,
  RawOoxmlNode,
  RawPackagePart,
} from "../source/index.js";

/** `writePptx` の出力。 */
export type WritePptxOutput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const textEncoder = new TextEncoder();

/**
 * CleanDoc source を PPTX package bytes に書き戻す。
 *
 * no-edit round-trip 用の初期 writer であり、未編集 package material を優先して
 * preserved output を作る。必要な raw bytes が無い non-bookkeeping part は、
 * 暗黙に再生成せずエラーにする。
 */
export function writePptx(source: CleanDocSource): WritePptxOutput {
  const files: Record<string, Uint8Array> = {
    [CONTENT_TYPES_PART]: encodeXml(serializeContentTypes(source.packageGraph.contentTypes)),
  };

  const written = new Set<string>([CONTENT_TYPES_PART]);

  for (const relationships of source.packageGraph.relationships) {
    const relsPath = relationshipsPartPath(relationships.sourcePartPath);
    files[relsPath] = encodeXml(serializeRelationships(relationships));
    written.add(relsPath);
  }

  for (const media of source.packageGraph.media) {
    files[media.partPath] = media.bytes;
    written.add(media.partPath);
  }

  for (const rawPart of source.packageGraph.rawParts ?? []) {
    files[rawPart.partPath] = serializeRawPackagePart(rawPart);
    written.add(rawPart.partPath);
  }

  for (const part of source.packageGraph.parts) {
    if (written.has(part.partPath)) continue;
    if (part.contentType === RELS_CONTENT_TYPE || isRelationshipPart(part.partPath)) continue;
    throw new Error(
      `writePptx: no preserved package material for part '${part.partPath}'; ` +
        "edited part generation is not implemented in the no-edit writer",
    );
  }

  return zipSync(files);
}

function serializeContentTypes(
  contentTypes: CleanDocSource["packageGraph"]["contentTypes"],
): string {
  const defaults = contentTypes.defaults
    .map(
      (entry) =>
        `<Default Extension="${escapeAttribute(entry.extension)}" ` +
        `ContentType="${escapeAttribute(entry.contentType)}"/>`,
    )
    .join("");
  const overrides = contentTypes.overrides
    .map(
      (entry) =>
        `<Override PartName="/${escapeAttribute(entry.partName)}" ` +
        `ContentType="${escapeAttribute(entry.contentType)}"/>`,
    )
    .join("");
  return (
    XML_DECLARATION +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    defaults +
    overrides +
    `</Types>`
  );
}

function serializeRelationships(partRelationships: PartRelationships): string {
  const relationships = partRelationships.relationships
    .map((relationship) => {
      const targetMode =
        relationship.targetMode === undefined
          ? ""
          : ` TargetMode="${escapeAttribute(relationship.targetMode)}"`;
      return (
        `<Relationship Id="${escapeAttribute(relationship.id)}" ` +
        `Type="${escapeAttribute(relationship.type)}" ` +
        `Target="${escapeAttribute(relationship.target)}"${targetMode}/>`
      );
    })
    .join("");
  return (
    XML_DECLARATION +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    relationships +
    `</Relationships>`
  );
}

function serializeRawPackagePart(rawPart: RawPackagePart): Uint8Array {
  if (rawPart.kind === "binary") return rawPart.bytes;
  return encodeXml(XML_DECLARATION + serializeRawNode(rawPart.xml));
}

function serializeRawNode(node: RawOoxmlNode): string {
  const attributes =
    node.attributes === undefined
      ? ""
      : Object.entries(node.attributes)
          .map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`)
          .join("");
  const text = node.text === undefined ? "" : escapeText(node.text);
  const children = node.children?.map((child) => serializeRawNode(child)).join("") ?? "";
  if (text === "" && children === "") return `<${node.name}${attributes}/>`;
  return `<${node.name}${attributes}>${text}${children}</${node.name}>`;
}

function relationshipsPartPath(sourcePartPath: PartPath): string {
  if (sourcePartPath === "") return "_rels/.rels";
  const slash = sourcePartPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : sourcePartPath.slice(0, slash + 1);
  const file = slash === -1 ? sourcePartPath : sourcePartPath.slice(slash + 1);
  return `${dir}_rels/${file}.rels`;
}

function isRelationshipPart(path: string): boolean {
  return path.endsWith(".rels") && path.includes("_rels/");
}

function encodeXml(xml: string): Uint8Array {
  return textEncoder.encode(xml);
}

function escapeAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeText(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}
