import { describe, it, expect } from "vitest";
import { parseXml } from "./xml-parser.js";

describe("parseXml", () => {
  it("parses simple XML", () => {
    const result = parseXml("<root><child>value</child></root>");
    expect((result.root as Record<string, unknown>).child).toBe("value");
  });

  it("preserves attributes with @_ prefix", () => {
    const result = parseXml('<shape type="rect" id="1"/>');
    const shape = result.shape as Record<string, unknown>;
    expect(shape["@_type"]).toBe("rect");
    expect(shape["@_id"]).toBe("1");
  });

  it("removes namespace prefix", () => {
    const result = parseXml('<p:sp xmlns:p="http://example.com"><p:nvSpPr/></p:sp>');
    expect(result).toHaveProperty("sp");
    const sp = (result.sp as Record<string, unknown>[])[0];
    expect(sp).toHaveProperty("nvSpPr");
  });
});

describe("ARRAY_TAGS", () => {
  const arrayTags = [
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
    "AlternateContent",
  ];

  it.each(arrayTags)("returns <%s> as array when single element", (tag) => {
    const xml = `<root><${tag}>content</${tag}></root>`;
    const result = parseXml(xml);
    const root = result.root as Record<string, unknown>;
    expect(Array.isArray(root[tag])).toBe(true);
    expect(root[tag]).toHaveLength(1);
  });

  it("returns sp as array when multiple elements", () => {
    const xml = "<spTree><sp><nvSpPr/></sp><sp><nvSpPr/></sp></spTree>";
    const result = parseXml(xml);
    const spTree = result.spTree as Record<string, unknown>;
    expect(Array.isArray(spTree.sp)).toBe(true);
    expect(spTree.sp).toHaveLength(2);
  });

  it("does not return non-array tag as array", () => {
    const xml = "<root><single>value</single></root>";
    const result = parseXml(xml);
    const root = result.root as Record<string, unknown>;
    expect(Array.isArray(root.single)).toBe(false);
    expect(root.single).toBe("value");
  });
});
