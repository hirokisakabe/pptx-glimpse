import { describe, expect, it } from "vitest";

import {
  asEmu,
  asPartPath,
  asRelationshipId,
  createComputedTemplateView,
  createPptx,
  createPptxAuthoringSession,
  type PptxSourceModel,
} from "../index.js";

describe("createComputedTemplateView", () => {
  it("projects a layout/master cascade without reading slide content or mutating source", () => {
    const initial = createPptx({
      slideMaster: {
        background: { kind: "solid", color: { kind: "srgb", hex: "AA0000" } },
      },
    });
    const master = requireValue(initial.slideMasters[0]);
    const layout = requireValue(initial.slideLayouts[0]);
    const slide = requireValue(initial.slides[0]);
    const session = createPptxAuthoringSession(initial);
    session.target(requireValue(master.handle)).addPlaceholder({
      type: "title",
      index: 4,
      name: "Master title",
      transform: {
        offsetX: asEmu(10),
        offsetY: asEmu(20),
        width: asEmu(300),
        height: asEmu(400),
      },
      promptText: "Master prompt",
    });
    session.target(requireValue(layout.handle)).addPlaceholder({
      type: "ctrTitle",
      index: 4,
      name: "Layout title",
      promptText: "Layout prompt",
    });
    session.target(requireValue(slide.handle)).addTextBox({
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      text: "Slide user content",
    });
    const source = session.source;
    const before = structuredClone(source);

    const computed = createComputedTemplateView(source, {
      kind: "slideLayout",
      partPath: layout.partPath,
    });

    expect(source).toEqual(before);
    expect(computed.slides).toHaveLength(1);
    expect(computed.slides[0]).toMatchObject({
      partPath: layout.partPath,
      layoutPartPath: layout.partPath,
      masterPartPath: master.partPath,
      background: { sourceLayer: "master", fill: { color: { hex: "#aa0000" } } },
    });
    const title = computed.slides[0]?.elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Layout title",
    );
    expect(title).toMatchObject({
      sourceLayer: "layout",
      transform: { offsetX: 10, offsetY: 20, width: 300, height: 400 },
      placeholderMatch: {
        layout: { name: "Layout title" },
        master: { name: "Master title" },
      },
    });
    expect(
      computed.slides[0]?.elements.some(
        (element) => element.sourceNode.name === "Slide user content",
      ),
    ).toBe(false);
  });

  it("resolves inherited master assets from their owning part when relationship IDs collide", () => {
    const source = buildTemplateAssetCollisionSource();
    const layout = requireValue(source.slideLayouts[0]);
    const computed = createComputedTemplateView(source, {
      kind: "slideLayout",
      partPath: layout.partPath,
    });
    const elements = requireValue(computed.slides[0]).elements;

    const masterImage = elements.find(
      (element) => element.kind === "image" && element.sourceNode.name === "Master image",
    );
    const layoutImage = elements.find(
      (element) => element.kind === "image" && element.sourceNode.name === "Layout image",
    );
    expect(masterImage?.kind).toBe("image");
    expect(layoutImage?.kind).toBe("image");
    if (masterImage?.kind !== "image" || layoutImage?.kind !== "image") {
      throw new Error("image collision fixture is incomplete");
    }
    expect(masterImage.relationship?.targetPartPath).toBe("ppt/media/master.png");
    expect(masterImage.media?.bytes).toEqual(new Uint8Array([1]));
    expect(layoutImage.relationship?.targetPartPath).toBe("ppt/media/layout.png");
    expect(layoutImage.media?.bytes).toEqual(new Uint8Array([2]));

    const masterChart = elements.find(
      (element) => element.kind === "chart" && element.sourceNode.name === "Master chart",
    );
    const layoutChart = elements.find(
      (element) => element.kind === "chart" && element.sourceNode.name === "Layout chart",
    );
    expect(masterChart?.kind).toBe("chart");
    expect(layoutChart?.kind).toBe("chart");
    if (masterChart?.kind !== "chart" || layoutChart?.kind !== "chart") {
      throw new Error("chart collision fixture is incomplete");
    }
    expect(masterChart.relationship?.targetPartPath).toBe("ppt/charts/master.xml");
    expect(masterChart.chartXml).toContain('owner="master"');
    expect(layoutChart.relationship?.targetPartPath).toBe("ppt/charts/layout.xml");
    expect(layoutChart.chartXml).toContain('owner="layout"');

    const masterSmartArt = elements.find(
      (element) => element.kind === "smartArt" && element.sourceNode.name === "Master SmartArt",
    );
    const layoutSmartArt = elements.find(
      (element) => element.kind === "smartArt" && element.sourceNode.name === "Layout SmartArt",
    );
    expect(masterSmartArt?.kind).toBe("smartArt");
    expect(layoutSmartArt?.kind).toBe("smartArt");
    if (masterSmartArt?.kind !== "smartArt" || layoutSmartArt?.kind !== "smartArt") {
      throw new Error("SmartArt collision fixture is incomplete");
    }
    expect(masterSmartArt.dataRelationship?.targetPartPath).toBe("ppt/diagrams/master-data.xml");
    expect(masterSmartArt.drawingPartPath).toBe("ppt/diagrams/master-drawing.xml");
    expect(masterSmartArt.drawingXml).toContain('owner="master"');
    expect(layoutSmartArt.dataRelationship?.targetPartPath).toBe("ppt/diagrams/layout-data.xml");
    expect(layoutSmartArt.drawingPartPath).toBe("ppt/diagrams/layout-drawing.xml");
    expect(layoutSmartArt.drawingXml).toContain('owner="layout"');
  });
});

function buildTemplateAssetCollisionSource(): PptxSourceModel {
  const source = createPptx();
  const master = requireValue(source.slideMasters[0]);
  const layout = requireValue(source.slideLayouts[0]);
  const imageRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const chartRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  const diagramDataRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
  const diagramDrawingRelationshipType =
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
  const ownerRelationships = (owner: "master" | "layout") => [
    {
      id: asRelationshipId("rIdImageCollision"),
      type: imageRelationshipType,
      target: `../media/${owner}.png`,
    },
    {
      id: asRelationshipId("rIdChartCollision"),
      type: chartRelationshipType,
      target: `../charts/${owner}.xml`,
    },
    {
      id: asRelationshipId("rIdSmartArtCollision"),
      type: diagramDataRelationshipType,
      target: `../diagrams/${owner}-data.xml`,
    },
  ];
  const ownerShapes = (owner: "Master" | "Layout") => [
    {
      kind: "image" as const,
      name: `${owner} image`,
      blipRelationshipId: asRelationshipId("rIdImageCollision"),
    },
    {
      kind: "chart" as const,
      name: `${owner} chart`,
      chartRelationshipId: asRelationshipId("rIdChartCollision"),
    },
    {
      kind: "smartArt" as const,
      name: `${owner} SmartArt`,
      dataRelationshipId: asRelationshipId("rIdSmartArtCollision"),
    },
  ];

  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      relationships: [
        ...source.packageGraph.relationships.map((entry) =>
          entry.sourcePartPath === master.partPath
            ? { ...entry, relationships: [...entry.relationships, ...ownerRelationships("master")] }
            : entry.sourcePartPath === layout.partPath
              ? {
                  ...entry,
                  relationships: [...entry.relationships, ...ownerRelationships("layout")],
                }
              : entry,
        ),
        ...(["master", "layout"] as const).map((owner) => ({
          sourcePartPath: asPartPath(`ppt/diagrams/${owner}-data.xml`),
          relationships: [
            {
              id: asRelationshipId("rIdDrawing"),
              type: diagramDrawingRelationshipType,
              target: `${owner}-drawing.xml`,
            },
          ],
        })),
      ],
      media: [
        ...source.packageGraph.media,
        {
          partPath: asPartPath("ppt/media/master.png"),
          contentType: "image/png",
          bytes: new Uint8Array([1]),
        },
        {
          partPath: asPartPath("ppt/media/layout.png"),
          contentType: "image/png",
          bytes: new Uint8Array([2]),
        },
      ],
      rawParts: [
        ...(source.packageGraph.rawParts ?? []),
        ...(["master", "layout"] as const).flatMap((owner) => [
          rawXmlPart(
            `ppt/charts/${owner}.xml`,
            `<c:chartSpace owner="${owner}"><c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>`,
          ),
          rawXmlPart(`ppt/diagrams/${owner}-data.xml`, "<dgm:dataModel/>"),
          rawXmlPart(
            `ppt/diagrams/${owner}-drawing.xml`,
            `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" owner="${owner}"/>`,
          ),
        ]),
      ],
    },
    slideMasters: source.slideMasters.map((candidate) =>
      candidate.partPath === master.partPath
        ? { ...candidate, shapes: ownerShapes("Master") }
        : candidate,
    ),
    slideLayouts: source.slideLayouts.map((candidate) =>
      candidate.partPath === layout.partPath
        ? { ...candidate, shapes: ownerShapes("Layout") }
        : candidate,
    ),
  };
}

function rawXmlPart(partPath: string, xml: string) {
  return {
    kind: "binary" as const,
    partPath: asPartPath(partPath),
    contentType: "application/xml",
    bytes: new TextEncoder().encode(xml),
  };
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("required fixture value is missing");
  return value;
}
