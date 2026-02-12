import { describe, it, expect, vi } from "vitest";
import { initWarningLogger } from "../warning-logger.js";
import {
  parseSlideMasterElements,
  parseSlideMasterColorMap,
  parseSlideMasterBackground,
  parseSlideMasterTxStyles,
} from "./slide-master-parser.js";
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

describe("parseSlideMasterTxStyles", () => {
  it("returns undefined when no txStyles", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld/>
      </p:sldMaster>
    `;
    expect(parseSlideMasterTxStyles(xml)).toBeUndefined();
  });

  it("parses titleStyle, bodyStyle, otherStyle", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld/>
        <p:txStyles>
          <p:titleStyle>
            <a:lvl1pPr algn="ctr">
              <a:defRPr sz="4400" b="1">
                <a:latin typeface="Calibri"/>
              </a:defRPr>
            </a:lvl1pPr>
          </p:titleStyle>
          <p:bodyStyle>
            <a:lvl1pPr marL="342900" indent="-342900" algn="l">
              <a:defRPr sz="3200"/>
            </a:lvl1pPr>
            <a:lvl2pPr marL="742950" indent="-285750" algn="l">
              <a:defRPr sz="2800"/>
            </a:lvl2pPr>
          </p:bodyStyle>
          <p:otherStyle>
            <a:defPPr>
              <a:defRPr sz="1800"/>
            </a:defPPr>
          </p:otherStyle>
        </p:txStyles>
      </p:sldMaster>
    `;
    const result = parseSlideMasterTxStyles(xml);

    expect(result).toBeDefined();

    // titleStyle
    expect(result!.titleStyle).toBeDefined();
    expect(result!.titleStyle!.levels[0]?.alignment).toBe("ctr");
    expect(result!.titleStyle!.levels[0]?.defaultRunProperties?.fontSize).toBe(44);
    expect(result!.titleStyle!.levels[0]?.defaultRunProperties?.bold).toBe(true);
    expect(result!.titleStyle!.levels[0]?.defaultRunProperties?.fontFamily).toBe("Calibri");

    // bodyStyle
    expect(result!.bodyStyle).toBeDefined();
    expect(result!.bodyStyle!.levels[0]?.marginLeft).toBe(342900);
    expect(result!.bodyStyle!.levels[0]?.indent).toBe(-342900);
    expect(result!.bodyStyle!.levels[0]?.defaultRunProperties?.fontSize).toBe(32);
    expect(result!.bodyStyle!.levels[1]?.marginLeft).toBe(742950);
    expect(result!.bodyStyle!.levels[1]?.defaultRunProperties?.fontSize).toBe(28);

    // otherStyle
    expect(result!.otherStyle).toBeDefined();
    expect(result!.otherStyle!.defaultParagraph?.defaultRunProperties?.fontSize).toBe(18);
  });

  it("returns undefined when txStyles is empty", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld/>
        <p:txStyles/>
      </p:sldMaster>
    `;
    expect(parseSlideMasterTxStyles(xml)).toBeUndefined();
  });

  it("parses partial txStyles (only titleStyle)", () => {
    const xml = `
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld/>
        <p:txStyles>
          <p:titleStyle>
            <a:lvl1pPr>
              <a:defRPr sz="4400"/>
            </a:lvl1pPr>
          </p:titleStyle>
        </p:txStyles>
      </p:sldMaster>
    `;
    const result = parseSlideMasterTxStyles(xml);

    expect(result).toBeDefined();
    expect(result!.titleStyle).toBeDefined();
    expect(result!.titleStyle!.levels[0]?.defaultRunProperties?.fontSize).toBe(44);
    expect(result!.bodyStyle).toBeUndefined();
    expect(result!.otherStyle).toBeUndefined();
  });

  it("warns and returns undefined for invalid XML", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterTxStyles(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('missing root element "sldMaster" in XML'),
    );
    expect(result).toBeUndefined();
    warnSpy.mockRestore();
    initWarningLogger("off");
  });
});

describe("structural validation warnings", () => {
  it("warns when parseSlideMasterColorMap receives XML without sldMaster root", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterColorMap(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('missing root element "sldMaster" in XML'),
    );
    expect(result.bg1).toBe("lt1");
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("warns when parseSlideMasterBackground receives XML without sldMaster root", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterBackground(xml, createColorResolver());

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('missing root element "sldMaster" in XML'),
    );
    expect(result).toBeNull();
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("warns when parseSlideMasterElements receives XML without sldMaster root", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseSlideMasterElements(
      xml,
      "ppt/slideMasters/slideMaster1.xml",
      createEmptyArchive(),
      createColorResolver(),
    );

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('missing root element "sldMaster" in XML'),
    );
    expect(result).toHaveLength(0);
    warnSpy.mockRestore();
    initWarningLogger("off");
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
