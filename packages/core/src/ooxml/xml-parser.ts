import { XMLParser } from "fast-xml-parser";

import { unsafeXmlBoundaryAssertion } from "../unsafe-type-assertion.js";

/** Internal note. */
export type XmlNode = Record<string, unknown>;

// Internal note.
// Internal note.
// Internal note.
// Internal note.
const ARRAY_TAGS = new Set([
  "sp", // Internal note.
  "pic", // Internal note.
  "cxnSp", // Internal note.
  "grpSp", // Internal note.
  "graphicFrame", // Internal note.
  "p", // Internal note.
  "r", // Internal note.
  "br", // Internal note.
  "fld", // Internal note.
  "Relationship", // Internal note.
  "sldId", // Internal note.
  "gs", // Internal note.
  "gridCol", // Internal note.
  "tr", // Internal note.
  "tc", // Internal note.
  "ser", // Internal note.
  "pt", // Internal note.
  "gd", // Internal note.
  "ds", // Internal note.
  "AlternateContent", // Internal note.
  "embeddedFont", // Internal note.
  "effectStyle", // Internal note.
  "font", // Internal note.
]);

// Internal note.
// Internal note.
const standardParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  htmlEntities: true,
  trimValues: false,
  isArray: (_name: string, jpath: unknown, _isLeafNode: boolean, _isAttribute: boolean) => {
    const tag = String(jpath).split(".").pop() ?? "";
    return ARRAY_TAGS.has(tag);
  },
});

export function parseXml(xml: string): Record<string, unknown> {
  return unsafeXmlBoundaryAssertion<Record<string, unknown>>(standardParser.parse(xml));
}
