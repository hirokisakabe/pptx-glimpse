import { XMLBuilder } from "fast-xml-parser";

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const textEncoder = new TextEncoder();
export const textDecoder = new TextDecoder();

export const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

export function encodeXml(xml: string): Uint8Array {
  return textEncoder.encode(xml);
}
