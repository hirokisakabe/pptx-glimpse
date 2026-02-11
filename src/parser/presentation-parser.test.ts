import { describe, it, expect, vi } from "vitest";
import { parsePresentation } from "./presentation-parser.js";

describe("parsePresentation", () => {
  it("parses valid presentation XML", () => {
    const xml = `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:sldIdLst>
          <p:sldId id="256" r:id="rId2"/>
          <p:sldId id="257" r:id="rId3"/>
        </p:sldIdLst>
      </p:presentation>
    `;
    const result = parsePresentation(xml);

    expect(result.slideSize.width).toBe(9144000);
    expect(result.slideSize.height).toBe(5143500);
    expect(result.slideRIds).toEqual(["rId2", "rId3"]);
  });

  it("warns and returns defaults when presentation root is missing", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `<other/>`;
    const result = parsePresentation(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('Presentation: missing root element "presentation"'),
    );
    expect(result.slideSize.width).toBe(9144000);
    expect(result.slideSize.height).toBe(5143500);
    expect(result.slideRIds).toEqual([]);

    warnSpy.mockRestore();
  });

  it("warns and uses default size when sldSz is missing", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:sldIdLst/>
      </p:presentation>
    `;
    const result = parsePresentation(xml);

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("Presentation: sldSz missing"));
    expect(result.slideSize.width).toBe(9144000);
    expect(result.slideSize.height).toBe(5143500);

    warnSpy.mockRestore();
  });

  it("does not warn for valid XML", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:sldIdLst>
          <p:sldId id="256" r:id="rId2"/>
        </p:sldIdLst>
      </p:presentation>
    `;
    parsePresentation(xml);

    expect(warnSpy).not.toHaveBeenCalled();

    warnSpy.mockRestore();
  });
});
