import { describe, expect, it } from "vitest";

import {
  asEmu,
  asOoxmlAngle,
  asPartPath,
  asPt,
  asRawSidecarId,
  asRelationshipId,
  asSourceNodeId,
  type CleanDocSource,
  type SourceImage,
  type SourceShape,
  type SourceTextRun,
} from "./index.js";

describe("CleanDoc source model types", () => {
  it("can build a minimal CleanDocSource value", () => {
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
      textBody: {
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
      handle: { partPath: slidePath, relationshipId: asRelationshipId("rId2") },
    };

    const source: CleanDocSource = {
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
        media: [{ partPath: mediaPath, contentType: "image/png", bytes: new Uint8Array([1, 2, 3]) }],
        rawParts: [
          {
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
      slides: [{ partPath: slidePath, layoutPartPath: layoutPath, shapes: [shape, image] }],
      slideLayouts: [{ partPath: layoutPath, masterPartPath: masterPath, shapes: [] }],
      slideMasters: [
        {
          partPath: masterPath,
          themePartPath: themePath,
          layoutPartPaths: [layoutPath],
          shapes: [],
        },
      ],
      themes: [{ partPath: themePath, name: "Office Theme" }],
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
    expect(source.presentation.slideSize?.width).toBe(9144000);
    expect(source.themes[0]?.partPath).toBe("ppt/theme/theme1.xml");
    expect(source.packageGraph.media[0]?.bytes).toEqual(new Uint8Array([1, 2, 3]));
  });
});
