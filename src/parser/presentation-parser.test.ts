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

  it("parses defaultTextStyle with multiple levels", () => {
    const xml = `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:sldIdLst>
          <p:sldId id="256" r:id="rId2"/>
        </p:sldIdLst>
        <p:defaultTextStyle>
          <a:defPPr>
            <a:defRPr lang="en-US"/>
          </a:defPPr>
          <a:lvl1pPr marL="0" algn="l">
            <a:defRPr sz="1800" b="1">
              <a:latin typeface="Calibri"/>
              <a:ea typeface="MS Gothic"/>
            </a:defRPr>
          </a:lvl1pPr>
          <a:lvl2pPr marL="457200" algn="l" indent="-228600">
            <a:defRPr sz="1400"/>
          </a:lvl2pPr>
        </p:defaultTextStyle>
      </p:presentation>
    `;
    const result = parsePresentation(xml);

    expect(result.defaultTextStyle).toBeDefined();
    expect(result.defaultTextStyle!.defaultParagraph).toBeUndefined();

    // lvl1pPr
    const lvl1 = result.defaultTextStyle!.levels[0];
    expect(lvl1).toBeDefined();
    expect(lvl1!.marginLeft).toBe(0);
    expect(lvl1!.alignment).toBe("l");
    expect(lvl1!.defaultRunProperties).toBeDefined();
    expect(lvl1!.defaultRunProperties!.fontSize).toBe(18);
    expect(lvl1!.defaultRunProperties!.bold).toBe(true);
    expect(lvl1!.defaultRunProperties!.fontFamily).toBe("Calibri");
    expect(lvl1!.defaultRunProperties!.fontFamilyEa).toBe("MS Gothic");

    // lvl2pPr
    const lvl2 = result.defaultTextStyle!.levels[1];
    expect(lvl2).toBeDefined();
    expect(lvl2!.marginLeft).toBe(457200);
    expect(lvl2!.indent).toBe(-228600);
    expect(lvl2!.defaultRunProperties!.fontSize).toBe(14);

    // lvl3pPr ~ lvl9pPr are undefined
    for (let i = 2; i < 9; i++) {
      expect(result.defaultTextStyle!.levels[i]).toBeUndefined();
    }
  });

  it("parses defaultTextStyle with defPPr containing defRPr", () => {
    const xml = `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:sldIdLst>
          <p:sldId id="256" r:id="rId2"/>
        </p:sldIdLst>
        <p:defaultTextStyle>
          <a:defPPr algn="ctr">
            <a:defRPr sz="2400" i="1" u="sng" strike="sngStrike"/>
          </a:defPPr>
        </p:defaultTextStyle>
      </p:presentation>
    `;
    const result = parsePresentation(xml);

    expect(result.defaultTextStyle).toBeDefined();
    const defPPr = result.defaultTextStyle!.defaultParagraph;
    expect(defPPr).toBeDefined();
    expect(defPPr!.alignment).toBe("ctr");
    expect(defPPr!.defaultRunProperties!.fontSize).toBe(24);
    expect(defPPr!.defaultRunProperties!.italic).toBe(true);
    expect(defPPr!.defaultRunProperties!.underline).toBe(true);
    expect(defPPr!.defaultRunProperties!.strikethrough).toBe(true);
  });

  it("returns undefined defaultTextStyle when element is missing", () => {
    const xml = `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:sldIdLst>
          <p:sldId id="256" r:id="rId2"/>
        </p:sldIdLst>
      </p:presentation>
    `;
    const result = parsePresentation(xml);

    expect(result.defaultTextStyle).toBeUndefined();
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
