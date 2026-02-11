import { describe, it, expect, vi } from "vitest";
import { parseFillFromNode } from "../../src/parser/fill-parser.js";
import type { FillParseContext } from "../../src/parser/fill-parser.js";
import { ColorResolver } from "../../src/color/color-resolver.js";
import type { PptxArchive } from "../../src/parser/pptx-reader.js";

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

describe("parseFillFromNode", () => {
  it("warns when gradFill has no gsLst", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const node = { gradFill: {} };
    const result = parseFillFromNode(node, createColorResolver());

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining("GradientFill: gsLst not found, skipping gradient"),
    );
    expect(result).toBeNull();
    warnSpy.mockRestore();
  });

  it("does not warn for valid gradient fill", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const node = {
      gradFill: {
        gsLst: {
          gs: [
            { "@_pos": "0", srgbClr: { "@_val": "FF0000" } },
            { "@_pos": "100000", srgbClr: { "@_val": "0000FF" } },
          ],
        },
      },
    };
    const result = parseFillFromNode(node, createColorResolver());

    expect(warnSpy).not.toHaveBeenCalled();
    expect(result?.type).toBe("gradient");
    warnSpy.mockRestore();
  });

  it("parses blipFill when context is provided", () => {
    const imageBuffer = Buffer.from("fake-png-data");
    const archive: PptxArchive = {
      files: new Map(),
      media: new Map([["ppt/media/image1.png", imageBuffer]]),
    };
    const context: FillParseContext = {
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
    const node = {
      blipFill: {
        blip: { "@_r:embed": "rId2" },
        stretch: { fillRect: {} },
      },
    };
    const result = parseFillFromNode(node, createColorResolver(), context);

    expect(result).not.toBeNull();
    expect(result?.type).toBe("image");
    if (result?.type === "image") {
      expect(result.mimeType).toBe("image/png");
      expect(result.imageData).toBe(imageBuffer.toString("base64"));
    }
  });

  it("ignores blipFill when context is not provided", () => {
    const node = {
      blipFill: {
        blip: { "@_r:embed": "rId2" },
      },
    };
    const result = parseFillFromNode(node, createColorResolver());
    expect(result).toBeNull();
  });

  it("returns null for blipFill with missing rel", () => {
    const archive: PptxArchive = {
      files: new Map(),
      media: new Map(),
    };
    const context: FillParseContext = {
      rels: new Map(),
      archive,
      basePath: "ppt/slideMasters/slideMaster1.xml",
    };
    const node = {
      blipFill: {
        blip: { "@_r:embed": "rId99" },
      },
    };
    const result = parseFillFromNode(node, createColorResolver(), context);
    expect(result).toBeNull();
  });
});
