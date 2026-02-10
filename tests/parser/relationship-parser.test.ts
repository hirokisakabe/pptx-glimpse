import { describe, it, expect } from "vitest";
import { buildRelsPath } from "../../src/parser/relationship-parser.js";

describe("buildRelsPath", () => {
  it("builds rels path for slide", () => {
    expect(buildRelsPath("ppt/slides/slide1.xml")).toBe("ppt/slides/_rels/slide1.xml.rels");
  });

  it("builds rels path for slide layout", () => {
    expect(buildRelsPath("ppt/slideLayouts/slideLayout1.xml")).toBe(
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    );
  });

  it("builds rels path for slide master", () => {
    expect(buildRelsPath("ppt/slideMasters/slideMaster1.xml")).toBe(
      "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    );
  });

  it("builds rels path for presentation", () => {
    expect(buildRelsPath("ppt/presentation.xml")).toBe("ppt/_rels/presentation.xml.rels");
  });
});
