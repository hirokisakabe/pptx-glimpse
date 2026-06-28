import { XMLParser } from "fast-xml-parser";

import { unsafeXmlBoundaryAssertion } from "../unsafe-type-assertion.js";

/** Type alias for XML nodes returned by fast-xml-parser */
export type XmlNode = Record<string, unknown>;

// Tags that require even a single element to be treated as an array in OOXML XML.
// fast-xml-parser returns an object if there is one child element, and an array if there are multiple child elements, so
// Parsing results become unstable when there is only one shape on the slide.
// By always using isArray to create an array, you can write downstream code in a unified manner.
const ARRAY_TAGS = new Set([
  "sp", // Shape
  "pic", // Picture
  "cxnSp", // Connector
  "grpSp", // Group (Group Shape)
  "graphicFrame", // Frames for tables, charts, etc.
  "p", // Text paragraph (Paragraph)
  "r", // Text run (Run)
  "br", // Line break (Break)
  "fld", // Field code (Field)
  "Relationship", // relationship
  "sldId", // Slide ID
  "gs", // Gradient Stop
  "gridCol", // table column definition
  "tr", // Table Row
  "tc", // Table Cell
  "ser", // Chart data series (Series)
  "pt", // Chart data point (Point)
  "gd", // Guide Definition
  "ds", // Custom Dash Segment
  "AlternateContent", // mc:AlternateContent (SmartArt etc.)
  "embeddedFont", // Embedded Font
  "effectStyle", // Effect Style
  "font", // Script-based Font definition
]);

// Singleton parser instance.
// XMLParser.parse() is stateless and can be safely reused.
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
