import { describe, it, expect, vi } from "vitest";
import { parseSlide, parseShapeTree } from "./slide-parser.js";
import { ColorResolver } from "../color/color-resolver.js";
import { parseXml } from "./xml-parser.js";
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

  it("parses normAutofit with fontScale and lnSpcReduction", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr>
              <a:normAutofit fontScale="62500" lnSpcReduction="20000"/>
            </a:bodyPr>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Hello</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    expect(shape.textBody).not.toBeNull();
    expect(shape.textBody!.bodyProperties.autoFit).toBe("normAutofit");
    expect(shape.textBody!.bodyProperties.fontScale).toBe(0.625);
    expect(shape.textBody!.bodyProperties.lnSpcReduction).toBe(0.2);
  });

  it("parses normAutofit without attributes as defaults", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr>
              <a:normAutofit/>
            </a:bodyPr>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Hello</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    expect(shape.textBody!.bodyProperties.autoFit).toBe("normAutofit");
    expect(shape.textBody!.bodyProperties.fontScale).toBe(1);
    expect(shape.textBody!.bodyProperties.lnSpcReduction).toBe(0);
  });

  it("defaults to noAutofit when no autofit element is present", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Hello</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    expect(shape.textBody!.bodyProperties.autoFit).toBe("noAutofit");
    expect(shape.textBody!.bodyProperties.fontScale).toBe(1);
    expect(shape.textBody!.bodyProperties.lnSpcReduction).toBe(0);
  });

  it("parses superscript baseline from rPr", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800" baseline="30000"/>
                <a:t>2</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    expect(shape.textBody!.paragraphs[0].runs[0].properties.baseline).toBe(30);
  });

  it("parses subscript baseline from rPr", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800" baseline="-25000"/>
                <a:t>2</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    expect(shape.textBody!.paragraphs[0].runs[0].properties.baseline).toBe(-25);
  });

  it("defaults baseline to 0 when not specified", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Normal</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    expect(shape.textBody!.paragraphs[0].runs[0].properties.baseline).toBe(0);
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

describe("bullet parsing", () => {
  it("parses buChar from paragraph properties", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:pPr lvl="0" marL="342900" indent="-342900">
                <a:buChar char="\u2022"/>
              </a:pPr>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Bullet item</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    const para = shape.textBody!.paragraphs[0];
    expect(para.properties.bullet).toEqual({ type: "char", char: "\u2022" });
    expect(para.properties.marginLeft).toBe(342900);
    expect(para.properties.indent).toBe(-342900);
  });

  it("parses buAutoNum from paragraph properties", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:pPr lvl="0">
                <a:buAutoNum type="arabicPeriod" startAt="3"/>
              </a:pPr>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Numbered item</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    const para = shape.textBody!.paragraphs[0];
    expect(para.properties.bullet).toEqual({
      type: "autoNum",
      scheme: "arabicPeriod",
      startAt: 3,
    });
  });

  it("parses buNone from paragraph properties", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:pPr>
                <a:buNone/>
              </a:pPr>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>No bullet</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    const para = shape.textBody!.paragraphs[0];
    expect(para.properties.bullet).toEqual({ type: "none" });
  });

  it("parses buFont from paragraph properties", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:pPr>
                <a:buFont typeface="Wingdings"/>
                <a:buChar char="v"/>
              </a:pPr>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Custom bullet</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    const para = shape.textBody!.paragraphs[0];
    expect(para.properties.bulletFont).toBe("Wingdings");
    expect(para.properties.bullet).toEqual({ type: "char", char: "v" });
  });

  it("defaults bullet to null when no bullet element is present", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="TextBox 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800"/>
                <a:t>Plain text</a:t>
              </a:r>
            </a:p>
          </p:txBody>
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

    const shape = elements[0] as ShapeElement;
    const para = shape.textBody!.paragraphs[0];
    expect(para.properties.bullet).toBeNull();
    expect(para.properties.marginLeft).toBe(0);
    expect(para.properties.indent).toBe(0);
  });
});

describe("structural validation warnings", () => {
  it("warns when parseSlide receives XML without sld root", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `<other><cSld/></other>`;
    const result = parseSlide(
      xml,
      "ppt/slides/slide3.xml",
      3,
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('Slide 3: missing root element "sld"'),
    );
    expect(result.elements).toEqual([]);

    warnSpy.mockRestore();
  });

  it("warns when a shape is skipped due to parse returning null", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Shape 1"/>
          </p:nvSpPr>
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
      "Slide 1",
    );

    expect(elements).toHaveLength(0);
    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("Slide 1: shape skipped"));

    warnSpy.mockRestore();
  });

  it("warns when NaN is detected in transform values", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Shape 1"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="abc" y="200"/>
              <a:ext cx="300" cy="400"/>
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
      "ppt/slides/slide2.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(1);
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining("NaN detected in transform offsetX"),
    );

    warnSpy.mockRestore();
  });

  it("does not warn for valid structures", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Shape 1"/>
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
    parseShapeTree(
      parsed.spTree,
      new Map(),
      "ppt/slides/slide1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).not.toHaveBeenCalled();

    warnSpy.mockRestore();
  });

  it("resolves theme font references (+mj-lt, +mn-lt) to actual font names", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Text 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US">
                  <a:latin typeface="+mj-lt"/>
                  <a:ea typeface="+mj-ea"/>
                </a:rPr>
                <a:t>Major</a:t>
              </a:r>
              <a:r>
                <a:rPr lang="en-US">
                  <a:latin typeface="+mn-lt"/>
                  <a:ea typeface="+mn-ea"/>
                </a:rPr>
                <a:t>Minor</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>
      </p:spTree>
    `;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = parseXml(xml) as any;
    const fontScheme = {
      majorFont: "Calibri Light",
      minorFont: "Calibri",
      majorFontEa: "Yu Gothic",
      minorFontEa: "Yu Mincho",
    };
    const elements = parseShapeTree(
      parsed.spTree,
      new Map(),
      "ppt/slides/slide1.xml",
      createEmptyArchive(),
      createColorResolver(),
      undefined,
      undefined,
      fontScheme,
    );

    expect(elements).toHaveLength(1);
    const shape = elements[0] as ShapeElement;
    const runs = shape.textBody!.paragraphs[0].runs;

    expect(runs[0].properties.fontFamily).toBe("Calibri Light");
    expect(runs[0].properties.fontFamilyEa).toBe("Yu Gothic");
    expect(runs[1].properties.fontFamily).toBe("Calibri");
    expect(runs[1].properties.fontFamilyEa).toBe("Yu Mincho");
  });

  it("does not resolve font names that are not theme references", () => {
    const xml = `
      <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Text 1"/>
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
          <p:txBody>
            <a:bodyPr/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US">
                  <a:latin typeface="Arial"/>
                  <a:ea typeface="Meiryo"/>
                </a:rPr>
                <a:t>Hello</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>
      </p:spTree>
    `;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = parseXml(xml) as any;
    const fontScheme = {
      majorFont: "Calibri Light",
      minorFont: "Calibri",
      majorFontEa: "Yu Gothic",
      minorFontEa: "Yu Mincho",
    };
    const elements = parseShapeTree(
      parsed.spTree,
      new Map(),
      "ppt/slides/slide1.xml",
      createEmptyArchive(),
      createColorResolver(),
      undefined,
      undefined,
      fontScheme,
    );

    expect(elements).toHaveLength(1);
    const shape = elements[0] as ShapeElement;
    const runs = shape.textBody!.paragraphs[0].runs;

    expect(runs[0].properties.fontFamily).toBe("Arial");
    expect(runs[0].properties.fontFamilyEa).toBe("Meiryo");
  });
});
