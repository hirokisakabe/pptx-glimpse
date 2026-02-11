import { describe, it, expect, vi } from "vitest";
import {
  parseSlideMasterElements,
  parseSlideMasterColorMap,
  parseSlideMasterBackground,
} from "../../src/parser/slide-master-parser.js";
import { ColorResolver } from "../../src/color/color-resolver.js";
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

describe("parseSlideMasterColorMap", () => {
  it("returns default color map when no clrMap", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld/>
      </p:sldMaster>
    `;
    const result = parseSlideMasterColorMap(xml);
    expect(result.bg1).toBe("lt1");
    expect(result.tx1).toBe("dk1");
  });
});

describe("parseSlideMasterBackground", () => {
  it("returns null when no background", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree/>
        </p:cSld>
      </p:sldMaster>
    `;
    const result = parseSlideMasterBackground(xml, createColorResolver());
    expect(result).toBeNull();
  });

  it("parses blipFill background when context is provided", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:cSld>
          <p:bg>
            <p:bgPr>
              <a:blipFill>
                <a:blip r:embed="rId2"/>
                <a:stretch><a:fillRect/></a:stretch>
              </a:blipFill>
            </p:bgPr>
          </p:bg>
          <p:spTree/>
        </p:cSld>
      </p:sldMaster>
    `;
    const imageBuffer = Buffer.from("test-image-data");
    const archive: PptxArchive = {
      files: new Map(),
      media: new Map([["ppt/media/image1.png", imageBuffer]]),
    };
    const context = {
      rels: new Map([
        [
          "rId2",
          {
            id: "rId2",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            target: "../media/image1.png",
          },
        ],
      ]),
      archive,
      basePath: "ppt/slideMasters/slideMaster1.xml",
    };
    const result = parseSlideMasterBackground(xml, createColorResolver(), context);
    expect(result).not.toBeNull();
    expect(result?.fill?.type).toBe("image");
    if (result?.fill?.type === "image") {
      expect(result.fill.mimeType).toBe("image/png");
      expect(result.fill.imageData).toBe(imageBuffer.toString("base64"));
    }
  });
});

describe("parseSlideMasterElements", () => {
  it("parses shapes from master", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
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
                <p:cNvPr id="5" name="Decorative"/>
                <p:cNvSpPr/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="9144000" cy="100"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
                <a:solidFill>
                  <a:srgbClr val="4472C4"/>
                </a:solidFill>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldMaster>
    `;

    const elements = parseSlideMasterElements(
      xml,
      "ppt/slideMasters/slideMaster1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(2);
    expect((elements[0] as ShapeElement).placeholderType).toBe("title");
    expect((elements[1] as ShapeElement).placeholderType).toBeUndefined();
    expect((elements[1] as ShapeElement).fill).toEqual({
      type: "solid",
      color: { hex: "#4472C4", alpha: 1 },
    });
  });

  it("returns empty array when no spTree", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld/>
      </p:sldMaster>
    `;

    const elements = parseSlideMasterElements(
      xml,
      "ppt/slideMasters/slideMaster1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(elements).toHaveLength(0);
  });
});

describe("structural validation warnings", () => {
  it("warns when parseSlideMasterColorMap receives XML without sldMaster root", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterColorMap(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('SlideMaster: missing root element "sldMaster"'),
    );
    expect(result.bg1).toBe("lt1");
    warnSpy.mockRestore();
  });

  it("warns when parseSlideMasterBackground receives XML without sldMaster root", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterBackground(xml, createColorResolver());

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('SlideMaster: missing root element "sldMaster"'),
    );
    expect(result).toBeNull();
    warnSpy.mockRestore();
  });

  it("warns when parseSlideMasterElements receives XML without sldMaster root", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterElements(
      xml,
      "ppt/slideMasters/slideMaster1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('SlideMaster: missing root element "sldMaster"'),
    );
    expect(result).toHaveLength(0);
    warnSpy.mockRestore();
  });

  it("does not warn for valid XML", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                  accent1="accent1" accent2="accent2" accent3="accent3"
                  accent4="accent4" accent5="accent5" accent6="accent6"
                  hlink="hlink" folHlink="folHlink"/>
        <p:cSld><p:spTree/></p:cSld>
      </p:sldMaster>
    `;
    parseSlideMasterColorMap(xml);
    parseSlideMasterBackground(xml, createColorResolver());
    parseSlideMasterElements(
      xml,
      "ppt/slideMasters/slideMaster1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).not.toHaveBeenCalled();
    warnSpy.mockRestore();
  });
});
