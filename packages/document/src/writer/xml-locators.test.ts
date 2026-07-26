import { describe, expect, it } from "vitest";

import { getChild, parseXml } from "../reader/xml.js";
import { locateShapeTreeNodeLocation } from "./xml-locators.js";

describe("shape tree node locator", () => {
  it("returns the unique nested node with its immediate group container and kind", () => {
    const spTree = parseShapeTree(
      shape("10", "Root") +
        group(
          "20",
          group("21", shape("22", "Nested target")) + connector("23", "Nested connector"),
        ),
    );
    const outer = getChild(spTree, "grpSp");
    const inner = getChild(outer, "grpSp");

    const location = locateShapeTreeNodeLocation(spTree, { nodeId: "22" });

    expect(location).toMatchObject({ nodeKind: "sp", nested: true });
    expect(location?.parentContainer).toBe(inner);
    expect(getChild(getChild(location?.node, "nvSpPr"), "cNvPr")?.["@_name"]).toBe("Nested target");
  });

  it("rejects duplicate ids across root and group descendants", () => {
    const spTree = parseShapeTree(shape("10", "Root") + group("20", shape("10", "Nested")));

    expect(() => locateShapeTreeNodeLocation(spTree, { nodeId: "10" })).toThrow(
      /duplicate shape id '10'/,
    );
  });

  it("rejects targets under mc:AlternateContent instead of choosing a branch", () => {
    const spTree = parseShapeTree(
      `<mc:AlternateContent>` +
        `<mc:Choice Requires="p14">${shape("30", "Choice")}</mc:Choice>` +
        `<mc:Fallback>${shape("31", "Fallback")}</mc:Fallback>` +
        `</mc:AlternateContent>`,
    );

    expect(() => locateShapeTreeNodeLocation(spTree, { nodeId: "31" })).toThrow(
      /inside mc:AlternateContent/,
    );
  });
});

function parseShapeTree(content: string) {
  const root = parseXml(
    `<p:sld xmlns:p="p" xmlns:a="a" xmlns:mc="mc">` +
      `<p:cSld><p:spTree>${content}</p:spTree></p:cSld>` +
      `</p:sld>`,
  );
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  if (spTree === undefined) throw new Error("spTree fixture was not parsed");
  return spTree;
}

function shape(id: string, name: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>`
  );
}

function connector(id: string, name: string): string {
  return (
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}"/>` +
    `<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>`
  );
}

function group(id: string, children: string): string {
  return (
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/>` +
    `<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${children}</p:grpSp>`
  );
}
