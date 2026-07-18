import { describe, expect, it } from "vitest";

import { getChild, getChildArray } from "../reader/xml.js";
import { replaceNodeEntries } from "./xml-node-utils.js";
import { buildXmlPreservingChildOrder, parseXmlForEditing } from "./xml-serialization.js";

describe("ordered XML editing serialization", () => {
  it("retains a replacement in its original heterogeneous sibling slot", () => {
    const root = parseXmlForEditing('<r><a id="1"/><b/><a id="2"/></r>');
    const element = getChild(root, "r")!;
    const second = getChildArray(element, "a")[1];

    replaceNodeEntries(element, [
      ["a", [{ "@_id": "new" }, second]],
      ["b", element.b],
    ]);

    expect(buildXmlPreservingChildOrder(root)).toBe('<r><a id="new"/><b/><a id="2"/></r>');
  });

  it("rejects non-whitespace mixed text that the grouped edit tree cannot represent", () => {
    expect(() => parseXmlForEditing("<r>pre<a/>mid<b/>post</r>")).toThrow(
      /mixed text\/element content/,
    );
  });

  it("does not duplicate formatting whitespace around ordered children", () => {
    const root = parseXmlForEditing("<r>\n  <a/>\n  <b/>\n</r>");
    const output = buildXmlPreservingChildOrder(root);

    expect(output.match(/\n/g)).toHaveLength(3);
    expect(output.indexOf("<a/>")).toBeLessThan(output.indexOf("<b/>"));
  });
});
