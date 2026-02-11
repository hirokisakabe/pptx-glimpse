import { XMLParser } from "fast-xml-parser";

const ARRAY_TAGS = new Set([
  "sp",
  "pic",
  "cxnSp",
  "grpSp",
  "graphicFrame",
  "p",
  "r",
  "br",
  "Relationship",
  "sldId",
  "gs",
  "gridCol",
  "tr",
  "tc",
  "ser",
  "pt",
  "gd",
]);

export function createXmlParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    removeNSPrefix: true,
    isArray: (_name: string, jpath: string) => {
      const tag = jpath.split(".").pop() ?? "";
      return ARRAY_TAGS.has(tag);
    },
  });
}

export function parseXml(xml: string): Record<string, unknown> {
  const parser = createXmlParser();
  return parser.parse(xml) as Record<string, unknown>;
}
