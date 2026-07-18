import {
  addChart,
  addConnector,
  addPicture,
  addShape,
  addTable,
  addTextBox,
  asEmu,
  createPptx,
  setSlideBackground,
  writePptx,
} from "@pptx-glimpse/document";

const CONTRACT_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x04, 0x08, 0x02, 0x00, 0x00, 0x00, 0x26, 0x93, 0x09,
  0x29, 0x00, 0x00, 0x00, 0x13, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x64, 0x60, 0xf8, 0xcf,
  0x00, 0x03, 0x4c, 0x70, 0x16, 0x5e, 0x0e, 0x00, 0x31, 0xd3, 0x01, 0x07, 0xd6, 0xce, 0xe7, 0x96,
  0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

/** Builds the committed authoring integration fixture through the document package's public API. */
export function createAuthoringIntegrationFixture(): Uint8Array {
  let source = createPptx({
    slideMaster: {
      name: "Authoring Contract Master",
      background: { kind: "solid", color: { kind: "srgb", hex: "F1F5F9" } },
    },
    slideLayout: {
      name: "Authoring Contract Layout",
      margin: {
        left: asEmu(120000),
        right: asEmu(120000),
        top: asEmu(80000),
        bottom: asEmu(80000),
      },
    },
  });

  const slide = source.slides[0];
  const master = source.slideMasters[0];
  const layout = source.slideLayouts[0];
  if (slide?.handle === undefined || master?.handle === undefined || layout?.handle === undefined) {
    throw new Error("createPptx did not create the expected slide, layout, and master handles");
  }

  source = addTextBox(source, master.handle, {
    offsetX: asEmu(300000),
    offsetY: asEmu(120000),
    width: asEmu(2600000),
    height: asEmu(300000),
    text: "MASTER CONTRACT",
    name: "Master contract text",
  });
  source = addTextBox(source, layout.handle, {
    offsetX: asEmu(6500000),
    offsetY: asEmu(120000),
    width: asEmu(2200000),
    height: asEmu(300000),
    text: "LAYOUT CONTRACT",
    name: "Layout contract text",
  });
  source = setSlideBackground(source, slide.handle, {
    kind: "solid",
    color: { kind: "srgb", hex: "FFF7ED" },
  });
  source = addTextBox(source, slide.handle, {
    offsetX: asEmu(450000),
    offsetY: asEmu(550000),
    width: asEmu(5000000),
    height: asEmu(550000),
    text: "Authoring integration fixture",
    name: "Fixture title",
  });
  const fixtureTitle = source.slides[0]?.shapes.find(
    (shapeNode) => shapeNode.kind === "shape" && shapeNode.name === "Fixture title",
  );
  if (fixtureTitle?.handle === undefined) {
    throw new Error("addTextBox did not create the expected source handle");
  }
  source = addShape(source, slide.handle, {
    geometry: { kind: "preset", preset: "roundRect" },
    offsetX: asEmu(600000),
    offsetY: asEmu(1400000),
    width: asEmu(2200000),
    height: asEmu(900000),
    fill: { kind: "solid", color: { kind: "srgb", hex: "2563EB" } },
    outline: {
      width: asEmu(18000),
      fill: { kind: "solid", color: { kind: "srgb", hex: "1E3A8A" } },
    },
    text: "Shape contract",
    name: "Contract shape",
  });

  const contractShape = source.slides[0]?.shapes.find(
    (shapeNode) => shapeNode.kind === "shape" && shapeNode.name === "Contract shape",
  );
  if (contractShape?.handle === undefined) {
    throw new Error("addShape did not create the expected source handle");
  }

  source = addPicture(source, slide.handle, {
    bytes: CONTRACT_PNG,
    offsetX: asEmu(3300000),
    offsetY: asEmu(1400000),
    width: asEmu(900000),
    height: asEmu(900000),
    name: "Contract picture",
  });
  source = addConnector(source, slide.handle, {
    preset: "straightConnector1",
    offsetX: asEmu(2800000),
    offsetY: asEmu(1800000),
    width: asEmu(500000),
    height: asEmu(100000),
    start: { shapeHandle: contractShape.handle, connectionSiteIndex: 1 },
    end: { shapeHandle: fixtureTitle.handle, connectionSiteIndex: 3 },
    name: "Contract connector",
    outline: { tailEnd: { type: "triangle", width: "sm", length: "sm" } },
  });
  source = addTable(source, slide.handle, {
    offsetX: asEmu(500000),
    offsetY: asEmu(2800000),
    width: asEmu(3800000),
    height: asEmu(1300000),
    columnWidths: [asEmu(1900000), asEmu(1900000)],
    rows: [
      {
        height: asEmu(650000),
        cells: [
          { text: "Table contract", fill: "DBEAFE" },
          { text: "Value", fill: "DBEAFE" },
        ],
      },
      { height: asEmu(650000), cells: [{ text: "Reader" }, { text: "Writer" }] },
    ],
    name: "Contract table",
  });
  source = addChart(source, slide.handle, {
    chartType: "bar",
    offsetX: asEmu(4700000),
    offsetY: asEmu(1200000),
    width: asEmu(3800000),
    height: asEmu(2900000),
    title: "Chart contract",
    showLegend: true,
    legendPosition: "b",
    series: [
      {
        name: "Coverage",
        categories: ["Reader", "Writer", "Renderer"],
        values: [3, 4, 5],
        color: "F97316",
      },
    ],
    name: "Contract chart",
  });

  return writePptx(source);
}
