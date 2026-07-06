/**
 * `writePptx(source)` - the first round-trip slice of the PptxSourceModel source writer.
 *
 * The writer targets structural round-trip preservation rather than byte equality. It is
 * not a package patcher that preserves XML attribute order, namespace prefix placement,
 * ZIP metadata, or defaulted OOXML values. Content types and relationships are
 * regenerated structurally from `packageGraph`; media bytes, unknown parts, and
 * non-bookkeeping raw parts prefer the raw package material preserved by the reader.
 * Only dirty scopes are updated according to supported PptxSourceModel operations.
 *
 * This file owns package write orchestration only. Dirty XML patching, presentation
 * topology patching, XML node helpers, edit validation, and part serializers live in
 * sibling modules so each writer responsibility has a single owner.
 */

import { zipSync } from "fflate";

import { editDirtyPartPath, editSlideTopologyOperation } from "../source/edit-descriptors.js";
import type { PptxSourceModel } from "../source/index.js";
import { isRelationshipPart, relationshipsPartPath } from "../source/package-paths.js";
import { serializeDirtyXmlPart } from "./dirty-part-edits.js";
import { validateEdits } from "./edit-validation.js";
import {
  serializeContentTypes,
  serializeRawPackagePart,
  serializeRelationships,
} from "./part-serializers.js";
import { serializePresentationWithSlideTopologyEdits } from "./presentation-topology.js";
import { encodeXml } from "./xml-serialization.js";

/** `writePptx` output. */
export type WritePptxOutput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

/**
 * Writes a PptxSourceModel source back to PPTX package bytes.
 *
 * This initial round-trip writer prefers unedited package material
 * to create preserved output. If the raw bytes needed to patch a dirty part are unavailable, or
 * a non-bookkeeping part lacks required raw bytes,
 * it throws instead of regenerating content implicitly.
 */
export function writePptx(source: PptxSourceModel): WritePptxOutput {
  const edits = source.edits ?? [];
  validateEdits(edits);
  const dirtyPartPaths = new Set(
    edits.flatMap((edit) => {
      const partPath = editDirtyPartPath(edit);
      return partPath === undefined ? [] : [partPath];
    }),
  );
  const slideTopologyOperations = edits.flatMap((edit) => {
    const operation = editSlideTopologyOperation(edit);
    return operation === undefined ? [] : [operation];
  });
  const hasSlideTopologyEdits = slideTopologyOperations.length > 0;
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
    if (hasSlideTopologyEdits && rawPart.partPath === source.presentation.partPath) continue;
    if (dirtyPartPaths.has(rawPart.partPath)) continue;
    files[rawPart.partPath] = serializeRawPackagePart(rawPart);
    written.add(rawPart.partPath);
  }

  for (const partPath of dirtyPartPaths) {
    files[partPath] = serializeDirtyXmlPart(source, partPath, edits);
    written.add(partPath);
  }

  if (hasSlideTopologyEdits) {
    files[source.presentation.partPath] = serializePresentationWithSlideTopologyEdits(
      source,
      slideTopologyOperations,
    );
    written.add(source.presentation.partPath);
  }

  for (const part of source.packageGraph.parts) {
    if (written.has(part.partPath)) continue;
    if (part.contentType === RELS_CONTENT_TYPE || isRelationshipPart(part.partPath)) continue;
    throw new Error(
      "writePptx: no preserved package material for part '" +
        part.partPath +
        "'; edited part generation is not implemented in the no-edit writer",
    );
  }

  return zipSync(files);
}
