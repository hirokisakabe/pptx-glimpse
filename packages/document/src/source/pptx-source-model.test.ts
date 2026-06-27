import { describe, expect, it } from "vitest";

// 実際の公開面 (`@pptx-glimpse/document/experimental`) 経由で import し、
// experimental entry point の re-export ごと検証する。
import {
  asEmu,
  asOoxmlAngle,
  asOoxmlPercent,
  asPartPath,
  asPt,
  asRawSidecarId,
  asRelationshipId,
  asSourceNodeId,
  type PptxSourceModel,
  type SourceImage,
  type SourceShape,
  type SourceTextRun,
} from "../experimental.js";

describe("PptxSourceModel source model types", () => {
  it("can build a minimal PptxSourceModel value", () => {
    const slidePath = asPartPath("ppt/slides/slide1.xml");
    const layoutPath = asPartPath("ppt/slideLayouts/slideLayout1.xml");
    const masterPath = asPartPath("ppt/slideMasters/slideMaster1.xml");
    const themePath = asPartPath("ppt/theme/theme1.xml");
    const mediaPath = asPartPath("ppt/media/image1.png");

    const shape: SourceShape = {
      kind: "shape",
      nodeId: asSourceNodeId("2"),
      name: "Title 1",
      transform: {
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(9144000),
        height: asEmu(1143000),
        rotation: asOoxmlAngle(0),
      },
      geometry: { preset: "rect" },
      fill: { kind: "solid", color: { kind: "scheme", scheme: "accent1" } },
      outline: {
        width: asEmu(12700),
        fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
      },
      textBody: {
        properties: { anchor: "middle", marginLeft: asEmu(91440) },
        paragraphs: [
          {
            properties: { align: "center" },
            runs: [
              {
                kind: "textRun",
                text: "Hello",
                properties: { bold: true, fontSize: asPt(44) },
              } satisfies SourceTextRun,
            ],
          },
        ],
      },
      handle: {
        partPath: slidePath,
        nodeId: asSourceNodeId("2"),
        orderingSlot: 0,
        rawSidecarIds: [asRawSidecarId("sidecar-1")],
      },
    };

    const image: SourceImage = {
      kind: "image",
      name: "Picture 1",
      blipRelationshipId: asRelationshipId("rId2"),
      crop: { left: asOoxmlPercent(10000), top: asOoxmlPercent(0) },
      handle: { partPath: slidePath, relationshipId: asRelationshipId("rId2") },
    };

    const source: PptxSourceModel = {
      packageGraph: {
        contentTypes: {
          defaults: [{ extension: "png", contentType: "image/png" }],
          overrides: [
            {
              partName: slidePath,
              contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            },
          ],
        },
        parts: [{ partPath: slidePath, contentType: "application/xml" }],
        relationships: [
          {
            sourcePartPath: slidePath,
            relationships: [
              {
                id: asRelationshipId("rId1"),
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                target: "../slideLayouts/slideLayout1.xml",
              },
            ],
          },
        ],
        media: [
          { partPath: mediaPath, contentType: "image/png", bytes: new Uint8Array([1, 2, 3]) },
        ],
        rawParts: [
          {
            kind: "xml",
            partPath: asPartPath("docProps/custom.xml"),
            contentType: "application/xml",
            xml: { name: "Properties" },
          },
        ],
      },
      presentation: {
        partPath: asPartPath("ppt/presentation.xml"),
        slideSize: { width: asEmu(9144000), height: asEmu(5143500) },
        slidePartPaths: [slidePath],
      },
      slides: [
        {
          partPath: slidePath,
          layoutPartPath: layoutPath,
          background: {
            kind: "fill",
            fill: { kind: "solid", color: { kind: "scheme", scheme: "bg1" } },
          },
          colorMapOverride: { mapping: { bg1: "lt1", tx1: "dk1" } },
          showMasterShapes: true,
          shapes: [shape, image],
        },
      ],
      slideLayouts: [
        { partPath: layoutPath, masterPartPath: masterPath, type: "title", shapes: [] },
      ],
      slideMasters: [
        {
          partPath: masterPath,
          themePartPath: themePath,
          layoutPartPaths: [layoutPath],
          colorMap: { mapping: { bg1: "lt1", tx1: "dk1" } },
          shapes: [],
        },
      ],
      themes: [
        {
          partPath: themePath,
          name: "Office Theme",
          colorScheme: {
            colors: { dk1: { kind: "system", value: "windowText", lastColor: "000000" } },
          },
          fontScheme: { majorLatin: "Calibri Light", minorLatin: "Calibri" },
        },
      ],
      diagnostics: [
        {
          severity: "warning",
          code: "unsupported-node-preserved",
          message: "kept raw extLst",
          handle: { partPath: slidePath },
        },
      ],
    };

    expect(source.slides[0]?.shapes).toHaveLength(2);
    expect(source.slides[0]?.showMasterShapes).toBe(true);
    expect(source.presentation.slideSize?.width).toBe(9144000);
    expect(source.themes[0]?.partPath).toBe("ppt/theme/theme1.xml");
    expect(source.themes[0]?.fontScheme?.minorLatin).toBe("Calibri");
    expect(source.packageGraph.media[0]?.bytes).toEqual(new Uint8Array([1, 2, 3]));
  });
});
