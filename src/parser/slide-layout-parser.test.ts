import { describe, it, expect, vi } from "vitest";
import { parseSlideLayoutBackground, parseSlideLayoutElements } from "./slide-layout-parser.js";
import { ColorResolver } from "../color/color-resolver.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { ShapeElement } from "../model/shape.js";

function createColorResolver() {
  return new ColorResolver(
    {
      dk1: "#000000",
      lt1: "#FFFFFF",
      dk2: "#44546A",
      lt2: "#E7E6E6",
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
      hlink: "#0563C1",
      folHlink: "#954F72",
    },
    {
      bg1: "lt1",
      tx1: "dk1",
      bg2: "lt2",
      tx2: "dk2",
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hlink: "hlink",
      folHlink: "folHlink",
    },
  );
}

function createEmptyArchive(): PptxArchive {
  return { files: new Map(), media: new Map() };
}

describe("parseSlideLayoutBackground", () => {
  it("returns null when no background defined", () => {
    const xml = `
      <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree/>
        </p:cSld>
      </p:sldLayout>
    `;
    const result = parseSlideLayoutBackground(xml, createColorResolver());
    expect(result).toBeNull();
  });
});

describe("parseSlideLayoutElements", () => {
  it("parses shapes from layout", () => {
    const xml = `
      <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="title"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="457200" y="274638"/>
                  <a:ext cx="8229600" cy="1143000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
              </p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="3" name="Footer"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="ftr" idx="10"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="3124200" y="6356350"/>
                  <a:ext cx="2895600" cy="365125"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `;

    const elements = parseSlideLayoutElements(
      xml,
      "ppt/slideLayouts/slideLayout1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(2);
    expect((elements[0] as ShapeElement).placeholderType).toBe("title");
    expect((elements[1] as ShapeElement).placeholderType).toBe("ftr");
    expect((elements[1] as ShapeElement).placeholderIdx).toBe(10);
  });

  it("returns empty array when no spTree", () => {
    const xml = `
      <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld/>
      </p:sldLayout>
    `;

    const elements = parseSlideLayoutElements(
      xml,
      "ppt/slideLayouts/slideLayout1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(0);
  });

  it("parses non-placeholder shapes from layout", () => {
    const xml = `
      <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="5" name="Logo"/>
                <p:cNvSpPr/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="100" y="200"/>
                  <a:ext cx="500" cy="500"/>
                </a:xfrm>
                <a:prstGeom prst="ellipse"/>
                <a:solidFill>
                  <a:srgbClr val="FF0000"/>
                </a:solidFill>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `;

    const elements = parseSlideLayoutElements(
      xml,
      "ppt/slideLayouts/slideLayout1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(1);
    const shape = elements[0] as ShapeElement;
    expect(shape.placeholderType).toBeUndefined();
    expect(shape.fill).toEqual({ type: "solid", color: { hex: "#FF0000", alpha: 1 } });
  });
});

describe("structural validation warnings", () => {
  it("warns when parseSlideLayoutBackground receives XML without sldLayout root", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideLayoutBackground(xml, createColorResolver());

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('SlideLayout: missing root element "sldLayout"'),
    );
    expect(result).toBeNull();
    warnSpy.mockRestore();
  });

  it("warns when parseSlideLayoutElements receives XML without sldLayout root", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideLayoutElements(
      xml,
      "ppt/slideLayouts/slideLayout1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('SlideLayout: missing root element "sldLayout"'),
    );
    expect(result).toHaveLength(0);
    warnSpy.mockRestore();
  });

  it("does not warn for valid XML", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld><p:spTree/></p:cSld>
      </p:sldLayout>
    `;
    parseSlideLayoutBackground(xml, createColorResolver());
    parseSlideLayoutElements(
      xml,
      "ppt/slideLayouts/slideLayout1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).not.toHaveBeenCalled();
    warnSpy.mockRestore();
  });
});
