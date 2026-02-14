import { describe, it, expect, vi } from "vitest";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { initWarningLogger } from "../warning-logger.js";
import type { FillParseContext } from "./fill-parser.js";
import { ColorResolver } from "../color/color-resolver.js";
import type { PptxArchive } from "./pptx-reader.js";

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
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const node = { gradFill: {} };
    const result = parseFillFromNode(node, createColorResolver());

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining("gsLst not found, skipping gradient"),
    );
    expect(result).toBeNull();
    warnSpy.mockRestore();
    initWarningLogger("off");
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
    if (result?.type === "gradient") {
      expect(result.gradientType).toBe("linear");
    }
    warnSpy.mockRestore();
  });

  it("parses radial gradient with fillToRect", () => {
    const node = {
      gradFill: {
        gsLst: {
          gs: [
            { "@_pos": "0", srgbClr: { "@_val": "FF0000" } },
            { "@_pos": "100000", srgbClr: { "@_val": "0000FF" } },
          ],
        },
        path: {
          "@_path": "circle",
          fillToRect: { "@_l": "50000", "@_t": "50000", "@_r": "50000", "@_b": "50000" },
        },
      },
    };
    const result = parseFillFromNode(node, createColorResolver());

    expect(result?.type).toBe("gradient");
    if (result?.type === "gradient") {
      expect(result.gradientType).toBe("radial");
      expect(result.centerX).toBe(0.5);
      expect(result.centerY).toBe(0.5);
      expect(result.stops).toHaveLength(2);
    }
  });

  it("parses radial gradient without fillToRect defaults to center", () => {
    const node = {
      gradFill: {
        gsLst: {
          gs: [
            { "@_pos": "0", srgbClr: { "@_val": "FF0000" } },
            { "@_pos": "100000", srgbClr: { "@_val": "0000FF" } },
          ],
        },
        path: { "@_path": "circle" },
      },
    };
    const result = parseFillFromNode(node, createColorResolver());

    expect(result?.type).toBe("gradient");
    if (result?.type === "gradient") {
      expect(result.gradientType).toBe("radial");
      expect(result.centerX).toBe(0.5);
      expect(result.centerY).toBe(0.5);
    }
  });

  it("parses radial gradient with top-left focus", () => {
    const node = {
      gradFill: {
        gsLst: {
          gs: [
            { "@_pos": "0", srgbClr: { "@_val": "FF0000" } },
            { "@_pos": "100000", srgbClr: { "@_val": "0000FF" } },
          ],
        },
        path: {
          "@_path": "circle",
          fillToRect: { "@_l": "0", "@_t": "0", "@_r": "100000", "@_b": "100000" },
        },
      },
    };
    const result = parseFillFromNode(node, createColorResolver());

    if (result?.type === "gradient") {
      expect(result.gradientType).toBe("radial");
      expect(result.centerX).toBe(0);
      expect(result.centerY).toBe(0);
    }
  });

  it("parses pattFill", () => {
    const node = {
      pattFill: {
        "@_prst": "ltDnDiag",
        fgClr: { srgbClr: { "@_val": "4472C4" } },
        bgClr: { srgbClr: { "@_val": "FFFFFF" } },
      },
    };
    const result = parseFillFromNode(node, createColorResolver());

    expect(result?.type).toBe("pattern");
    if (result?.type === "pattern") {
      expect(result.preset).toBe("ltDnDiag");
      expect(result.foregroundColor.hex).toBe("#4472C4");
      expect(result.backgroundColor.hex).toBe("#FFFFFF");
    }
  });

  it("returns null for pattFill without colors", () => {
    const node = {
      pattFill: {
        "@_prst": "ltDnDiag",
      },
    };
    const result = parseFillFromNode(node, createColorResolver());
    expect(result).toBeNull();
  });

  it("parses grpFill using context groupFill", () => {
    const context: FillParseContext = {
      rels: new Map(),
      archive: { files: new Map(), media: new Map() },
      basePath: "ppt/slides/slide1.xml",
      groupFill: { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
    };
    const node = { grpFill: {} };
    const result = parseFillFromNode(node, createColorResolver(), context);

    expect(result?.type).toBe("solid");
    if (result?.type === "solid") {
      expect(result.color.hex).toBe("#FF0000");
    }
  });

  it("returns null for grpFill without context groupFill", () => {
    const node = { grpFill: {} };
    const result = parseFillFromNode(node, createColorResolver());
    expect(result).toBeNull();
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

  it("parses blipFill with tile attributes", () => {
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
      basePath: "ppt/slides/slide1.xml",
    };
    const node = {
      blipFill: {
        blip: { "@_r:embed": "rId2" },
        tile: { "@_sx": "50000", "@_sy": "50000", "@_tx": "0", "@_ty": "0", "@_flip": "x" },
      },
    };
    const result = parseFillFromNode(node, createColorResolver(), context);

    expect(result).not.toBeNull();
    expect(result?.type).toBe("image");
    if (result?.type === "image") {
      expect(result.tile).not.toBeNull();
      expect(result.tile!.sx).toBe(0.5);
      expect(result.tile!.sy).toBe(0.5);
      expect(result.tile!.flip).toBe("x");
    }
  });

  it("parses blipFill without tile returns null tile", () => {
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
      basePath: "ppt/slides/slide1.xml",
    };
    const node = {
      blipFill: {
        blip: { "@_r:embed": "rId2" },
        stretch: { fillRect: {} },
      },
    };
    const result = parseFillFromNode(node, createColorResolver(), context);

    expect(result?.type).toBe("image");
    if (result?.type === "image") {
      expect(result.tile).toBeNull();
    }
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

describe("parseOutline", () => {
  it("parses gradient fill in outline", () => {
    const lnNode = {
      "@_w": "25400",
      gradFill: {
        gsLst: {
          gs: [
            { "@_pos": "0", srgbClr: { "@_val": "FF0000" } },
            { "@_pos": "100000", srgbClr: { "@_val": "0000FF" } },
          ],
        },
        lin: { "@_ang": "5400000" },
      },
    };
    const result = parseOutline(lnNode, createColorResolver());
    expect(result).not.toBeNull();
    expect(result!.fill?.type).toBe("gradient");
    if (result!.fill?.type === "gradient") {
      expect(result!.fill.gradientType).toBe("linear");
      expect(result!.fill.angle).toBe(90);
      expect(result!.fill.stops).toHaveLength(2);
    }
  });

  it("parses custDash in outline", () => {
    const lnNode = {
      "@_w": "25400",
      solidFill: { srgbClr: { "@_val": "000000" } },
      custDash: {
        ds: [
          { "@_d": "300000", "@_sp": "100000" },
          { "@_d": "100000", "@_sp": "100000" },
        ],
      },
    };
    const result = parseOutline(lnNode, createColorResolver());
    expect(result).not.toBeNull();
    expect(result!.customDash).toEqual([3, 1, 1, 1]);
  });

  it("returns undefined customDash when no custDash element", () => {
    const lnNode = {
      "@_w": "25400",
      solidFill: { srgbClr: { "@_val": "000000" } },
    };
    const result = parseOutline(lnNode, createColorResolver());
    expect(result).not.toBeNull();
    expect(result!.customDash).toBeUndefined();
  });
});
