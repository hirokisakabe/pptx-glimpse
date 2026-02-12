import { describe, it, expect, vi } from "vitest";
import { buildRelsPath, parseRelationships } from "./relationship-parser.js";
import { initWarningLogger } from "../warning-logger.js";

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

describe("parseRelationships", () => {
  it("parses valid relationships XML", () => {
    const xml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
      </Relationships>
    `;
    const result = parseRelationships(xml);

    expect(result.size).toBe(2);
    expect(result.get("rId1")?.target).toBe("../slideLayouts/slideLayout1.xml");
    expect(result.get("rId2")?.target).toBe("../theme/theme1.xml");
  });

  it("warns when Relationships root is missing", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseRelationships(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('missing root element "Relationships" in XML'),
    );
    expect(result.size).toBe(0);
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("warns and skips entries missing required attributes", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://example.com/type" Target="target.xml"/>
        <Relationship Type="http://example.com/type" Target="target2.xml"/>
        <Relationship Id="rId3" Target="target3.xml"/>
        <Relationship Id="rId4" Type="http://example.com/type"/>
      </Relationships>
    `;
    const result = parseRelationships(xml);

    expect(warnSpy).toHaveBeenCalledTimes(3);
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining("entry missing required attribute, skipping"),
    );
    expect(result.size).toBe(1);
    expect(result.get("rId1")?.target).toBe("target.xml");
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("does not warn for valid XML", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://example.com/type" Target="target.xml"/>
      </Relationships>
    `;
    parseRelationships(xml);

    expect(warnSpy).not.toHaveBeenCalled();
    warnSpy.mockRestore();
  });
});
