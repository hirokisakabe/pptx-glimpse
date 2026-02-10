import { describe, it, expect } from "vitest";
import { parseShapeTree } from "../../src/parser/slide-parser.js";
import { ColorResolver } from "../../src/color/color-resolver.js";
import { parseXml } from "../../src/parser/xml-parser.js";
import type { PptxArchive } from "../../src/parser/pptx-reader.js";
import type { ShapeElement } from "../../src/model/shape.js";

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

describe("parseShapeTree", () => {
  it("parses placeholder type from shape", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Title 1"/>
            <p:cNvSpPr/>
            <p:nvPr>
              <p:ph type="title"/>
            </p:nvPr>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="100" y="200"/>
              <a:ext cx="300" cy="400"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
          </p:spPr>
        </p:sp>
      </p:spTree>
    `;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = parseXml(xml) as any;
    const spTree = parsed.spTree;
    const elements = parseShapeTree(
      spTree,
      new Map(),
      "ppt/slides/slide1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(1);
    expect(elements[0].type).toBe("shape");
    expect((elements[0] as ShapeElement).placeholderType).toBe("title");
  });

  it("defaults placeholder type to body when type is not specified", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="3" name="Content 1"/>
            <p:cNvSpPr/>
            <p:nvPr>
              <p:ph idx="1"/>
            </p:nvPr>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="100" y="200"/>
              <a:ext cx="300" cy="400"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
          </p:spPr>
        </p:sp>
      </p:spTree>
    `;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = parseXml(xml) as any;
    const spTree = parsed.spTree;
    const elements = parseShapeTree(
      spTree,
      new Map(),
      "ppt/slides/slide1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(1);
    const shape = elements[0] as ShapeElement;
    expect(shape.placeholderType).toBe("body");
    expect(shape.placeholderIdx).toBe(1);
  });

  it("does not set placeholder fields for non-placeholder shapes", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="4" name="Rectangle 1"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="100" y="200"/>
              <a:ext cx="300" cy="400"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
          </p:spPr>
        </p:sp>
      </p:spTree>
    `;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = parseXml(xml) as any;
    const spTree = parsed.spTree;
    const elements = parseShapeTree(
      spTree,
      new Map(),
      "ppt/slides/slide1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(1);
    const shape = elements[0] as ShapeElement;
    expect(shape.placeholderType).toBeUndefined();
    expect(shape.placeholderIdx).toBeUndefined();
  });

  it("parses various placeholder types", () => {
    const types = ["title", "body", "dt", "ftr", "sldNum", "ctrTitle", "subTitle"];
    for (const phType of types) {
      const xml = `
        <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Shape"/>
              <p:cNvSpPr/>
              <p:nvPr>
                <p:ph type="${phType}"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="100" cy="100"/>
              </a:xfrm>
              <a:prstGeom prst="rect"/>
            </p:spPr>
          </p:sp>
        </p:spTree>
      `;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const parsed = parseXml(xml) as any;
      const elements = parseShapeTree(
        parsed.spTree,
        new Map(),
        "ppt/slides/slide1.xml",
        createEmptyArchive(),
        createColorResolver(),
      );
      expect((elements[0] as ShapeElement).placeholderType).toBe(phType);
    }
  });
});
