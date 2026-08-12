import { Buffer } from "node:buffer";

import {
  addChart,
  addConnector,
  addPicture,
  addShape,
  addTable,
  asEmu,
  createPptx,
  groupShapes,
  type PptxSourceModel,
  readPptx,
  type SourceChart,
  type SourceConnector,
  type SourceHandle,
  type SourceImage,
  type SourceShape,
  type SourceShapeNode,
  writePptx,
} from "@pptx-glimpse/document";
import JSZip from "jszip";

import { type EditorApplyCommandResult, type EditorHistoryResult } from "./index.js";

const encoder = new TextEncoder();

export const RED_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=",
);

export const BLUE_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
);

export const JPEG_BYTES = new Uint8Array([0xff, 0xd8, 0xff, 0xd9]);

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function pngBytes(base64: string): Uint8Array {
  return new Uint8Array(Buffer.from(base64, "base64"));
}

export function createThreeShapeSource(): PptxSourceModel {
  let source = createPptx();
  const slideHandle = requireHandle(source.slides[0]?.handle);
  for (const offsetX of [0, 2000, 4000]) {
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(offsetX),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
    });
  }
  return source;
}

export function buildDrawingDeleteFixture(): Uint8Array {
  let source = createPptx();
  const slideHandle = requireHandle(source.slides[0]?.handle);
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(100),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Delete Shape",
  });
  source = addPicture(source, slideHandle, {
    bytes: RED_PNG,
    offsetX: asEmu(1200),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Delete Picture",
  });
  source = addTable(source, slideHandle, {
    offsetX: asEmu(2300),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    columnWidths: [asEmu(1000)],
    rows: [{ height: asEmu(1000), cells: [{ text: "Cell" }] }],
    name: "Delete Table",
  });
  source = addChart(source, slideHandle, {
    chartType: "bar",
    series: [{ categories: ["A"], values: [1] }],
    offsetX: asEmu(3400),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Delete Chart",
  });
  for (const [name, offsetX] of [
    ["Group A", 4500],
    ["Group B", 5600],
  ] as const) {
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(offsetX),
      offsetY: asEmu(100),
      width: asEmu(1000),
      height: asEmu(1000),
      name,
    });
  }
  source = groupShapes(
    source,
    source.slides[0].shapes
      .filter((shape) => shape.kind !== "raw" && shape.name?.startsWith("Group "))
      .map((shape) => requireHandle(shape.handle)),
  );
  return writePptx(source);
}

export function buildNestedDrawingDeleteFixture(): Uint8Array {
  let source = readPptx(buildDrawingDeleteFixture());
  const slideHandle = requireHandle(source.slides[0]?.handle);
  source = addConnector(source, slideHandle, {
    preset: "straightConnector1",
    offsetX: asEmu(6700),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Delete Connector",
  });
  source = groupShapes(
    source,
    source.slides[0].shapes.map((shape) => requireHandle(shape.handle)),
  );
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(7800),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Outer Sibling",
  });
  source = groupShapes(
    source,
    source.slides[0].shapes.map((shape) => requireHandle(shape.handle)),
  );
  return writePptx(source);
}

export async function buildTextEditFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:pPr algn="ctr"/>` +
        `<a:r><a:rPr b="1" sz="2400"><a:latin typeface="Aptos"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Original</a:t></a:r>` +
        `<a:r><a:rPr i="1" sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t xml:space="preserve"> Keep </a:t></a:r>` +
        `</a:p>` +
        `</p:txBody></p:sp>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="11" name="No xfrm"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>No xfrm</a:t></a:r></a:p></p:txBody></p:sp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

export async function buildShapeStyleFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Styled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Styled</a:t></a:r></a:p></p:txBody></p:sp>` +
        `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="20" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
        `<p:spPr><a:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></a:xfrm>` +
        `<a:prstGeom prst="straightConnector1"/><a:ln w="12700"><a:noFill/></a:ln></p:spPr>` +
        `</p:cxnSp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

export async function buildImageReplacementFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Shared Picture A"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="21" name="Shared Picture B"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/media/image1.png", RED_PNG);

  return zip.generateAsync({ type: "uint8array" });
}

export async function buildTwoSlideSharedImageFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/slides/slide1.xml", imageSlideXml(20, "Shared Picture A"));
  zip.file("ppt/slides/slide2.xml", imageSlideXml(30, "Shared Picture B"));
  for (const slideIndex of [1, 2]) {
    zip.file(
      `ppt/slides/_rels/slide${slideIndex}.xml.rels`,
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
          `</Relationships>`,
      ),
    );
  }
  zip.file("ppt/media/image1.png", RED_PNG);

  return zip.generateAsync({ type: "uint8array" });
}

function imageSlideXml(shapeId: number, name: string): Uint8Array {
  return xml(
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<p:cSld><p:spTree>` +
      `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
      `</p:spTree></p:cSld>` +
      `</p:sld>`,
  );
}

export async function buildTwoSlideFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="10" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );
  zip.file(
    "ppt/slides/slide2.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="20" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

export async function buildUnreferencedLayoutFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sld>`,
    ),
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">` +
        `<p:cSld name="Referenced"><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slideLayouts/slideLayout2.xml",
    xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">` +
        `<p:cSld name="Unreferenced"><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `<p:sldLayoutIdLst>` +
        `<p:sldLayoutId id="2147483649" r:id="rIdLayout1"/>` +
        `<p:sldLayoutId id="2147483650" r:id="rIdLayout2"/>` +
        `</p:sldLayoutIdLst>` +
        `</p:sldMaster>`,
    ),
  );
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `<Relationship Id="rIdLayout2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>` +
        `</Relationships>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

export function expectApplied(result: EditorApplyCommandResult): PptxSourceModel {
  if (!result.ok) throw new Error(result.message);
  return result.document;
}

export function expectHistory(result: EditorHistoryResult): PptxSourceModel {
  if (!result.ok) throw new Error(result.message);
  return result.document;
}

export function firstShape(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0].shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("shape not found");
  return shape;
}

export function firstParagraph(source: PptxSourceModel) {
  return firstShape(source).textBody!.paragraphs[0];
}

export function firstRun(source: PptxSourceModel) {
  return firstParagraph(source).runs[0];
}

export function firstImage(source: PptxSourceModel): SourceImage {
  const image = source.slides[0].shapes.find((node): node is SourceImage => node.kind === "image");
  if (image === undefined) throw new Error("image not found");
  return image;
}

export function firstChart(source: PptxSourceModel): SourceChart {
  const chart = source.slides[0].shapes.find((node): node is SourceChart => node.kind === "chart");
  if (chart === undefined) throw new Error("chart not found");
  return chart;
}

export function buildChartEditSource(): PptxSourceModel {
  let source = createPptx();
  const slideHandle = requireHandle(source.slides[0]?.handle);
  source = addChart(source, slideHandle, {
    chartType: "bar",
    offsetX: asEmu(100),
    offsetY: asEmu(200),
    width: asEmu(3000),
    height: asEmu(2000),
    series: [
      { name: "Original 1", categories: ["A", "B"], values: [1, 2] },
      { name: "Original 2", categories: ["A", "B"], values: [3, 4] },
    ],
  });
  return readPptx(writePptx(source));
}

export async function buildCategoryComboChartEditSource(
  options: {
    readonly lineGrouping?: string;
    readonly lineAxisIds?: readonly (string | undefined)[];
  } = {},
): Promise<PptxSourceModel> {
  const pptx = await JSZip.loadAsync(writePptx(buildChartEditSource()));
  const chartFile = pptx.file("ppt/charts/chart1.xml");
  if (chartFile === null) throw new Error("combo chart fixture part not found");
  let chart = await chartFile.async("string");
  const barChart = /<c:barChart>.*?<\/c:barChart>/.exec(chart)?.[0];
  const series = barChart?.match(/<c:ser>.*?<\/c:ser>/g);
  if (barChart === undefined || series?.length !== 2) {
    throw new Error("combo chart fixture requires two category series");
  }
  const lineSeries = series[1].replace(
    '<c:idx val="1"/><c:order val="1"/>',
    '<c:idx val="7"/><c:order val="9"/>',
  );
  const lineAxisIds = (options.lineAxisIds ?? ["100002", "100003"])
    .map((id) => (id === undefined ? "<c:axId/>" : `<c:axId val="${id}"/>`))
    .join("");
  const lineChart = `<c:lineChart><c:grouping val="${options.lineGrouping ?? "standard"}"/><c:varyColors val="0"/>${lineSeries}${lineAxisIds}</c:lineChart>`;
  chart = chart.replace(barChart, `${barChart.replace(series[1], "")}${lineChart}`);
  pptx.file("ppt/charts/chart1.xml", chart);
  return readPptx(await pptx.generateAsync({ type: "uint8array" }));
}

export async function buildScatterChartEditSource(): Promise<PptxSourceModel> {
  const pptx = await JSZip.loadAsync(writePptx(buildChartEditSource()));
  const chartFile = pptx.file("ppt/charts/chart1.xml");
  const workbookFile = pptx.file("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx");
  if (chartFile === null || workbookFile === null) throw new Error("chart fixture parts not found");
  const chart = await chartFile.async("string");
  const scatterSeries = (
    index: number,
    name: string,
    headerRow: number,
    x: number[],
    y: number[],
  ) => {
    const lastRow = headerRow + x.length;
    const points = (values: number[]) =>
      values.map((value, point) => `<c:pt idx="${point}"><c:v>${value}</c:v></c:pt>`).join("");
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$B$${headerRow}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:xVal><c:numRef><c:f>Sheet1!$A$${headerRow + 1}:$A$${lastRow}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${x.length}"/>${points(x)}</c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Sheet1!$B$${headerRow + 1}:$B$${lastRow}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${y.length}"/>${points(y)}</c:numCache></c:numRef></c:yVal></c:ser>`;
  };
  pptx.file(
    "ppt/charts/chart1.xml",
    replaceCategoryAxisWithValueAxis(
      chart.replace(
        /<c:barChart>.*?<\/c:barChart>/,
        `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${scatterSeries(0, "Original 1", 1, [1, 2], [1, 2])}${scatterSeries(1, "Original 2", 5, [3, 4], [3, 4])}<c:axId val="100002"/><c:axId val="100003"/></c:scatterChart>`,
      ),
    ),
  );
  const workbook = await JSZip.loadAsync(await workbookFile.async("uint8array"));
  workbook.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Original 1</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>2</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>Original 2</t></is></c></row><row r="6"><c r="A6"><v>3</v></c><c r="B6"><v>3</v></c></row><row r="7"><c r="A7"><v>4</v></c><c r="B7"><v>4</v></c></row></sheetData></worksheet>`,
  );
  pptx.file(
    "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    await workbook.generateAsync({ type: "uint8array" }),
  );
  return readPptx(await pptx.generateAsync({ type: "uint8array" }));
}

export async function buildBubbleChartEditSource(): Promise<PptxSourceModel> {
  const pptx = await JSZip.loadAsync(writePptx(await buildScatterChartEditSource()));
  const chartFile = pptx.file("ppt/charts/chart1.xml");
  const workbookFile = pptx.file("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx");
  if (chartFile === null || workbookFile === null)
    throw new Error("bubble fixture parts not found");
  let chart = (await chartFile.async("string"))
    .replace(
      '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>',
      '<c:bubbleChart><c:varyColors val="0"/>',
    )
    .replace("</c:scatterChart>", '<c:bubbleScale val="100"/></c:bubbleChart>');
  for (const [range, values] of [
    ["2:$C$3", [4, 8]],
    ["6:$C$7", [6, 9]],
  ] as const) {
    const points = values
      .map((value, point) => `<c:pt idx="${point}"><c:v>${value}</c:v></c:pt>`)
      .join("");
    chart = chart.replace(
      "</c:yVal></c:ser>",
      `</c:yVal><c:bubbleSize><c:numRef><c:f>Sheet1!$C$${range}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef></c:bubbleSize></c:ser>`,
    );
  }
  pptx.file("ppt/charts/chart1.xml", chart);
  const workbook = await JSZip.loadAsync(await workbookFile.async("uint8array"));
  workbook.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Original 1</t></is></c><c r="C1" t="inlineStr"><is><t>Size</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c><c r="C2"><v>4</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>2</v></c><c r="C3"><v>8</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>Original 2</t></is></c><c r="C5" t="inlineStr"><is><t>Size</t></is></c></row><row r="6"><c r="A6"><v>3</v></c><c r="B6"><v>3</v></c><c r="C6"><v>6</v></c></row><row r="7"><c r="A7"><v>4</v></c><c r="B7"><v>4</v></c><c r="C7"><v>9</v></c></row></sheetData></worksheet>`,
  );
  pptx.file(
    "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    await workbook.generateAsync({ type: "uint8array" }),
  );
  return readPptx(await pptx.generateAsync({ type: "uint8array" }));
}

function replaceCategoryAxisWithValueAxis(chartXml: string): string {
  const categoryAxis = /<c:catAx>.*?<\/c:catAx>/.exec(chartXml)?.[0];
  if (categoryAxis === undefined) throw new Error("fixture category axis not found");
  const valueAxis = categoryAxis
    .replace("<c:catAx>", "<c:valAx>")
    .replace("</c:catAx>", '<c:crossBetween val="midCat"/></c:valAx>')
    .replace(/<c:(?:auto|lblAlgn|lblOffset|noMultiLvlLbl)\b[^>]*\/>/g, "");
  return chartXml.replace(categoryAxis, valueAxis);
}

export function chartXml(source: PptxSourceModel): string {
  const chartPart = source.packageGraph.rawParts?.find(
    (part) => part.partPath === "ppt/charts/chart1.xml",
  );
  if (chartPart?.kind !== "binary") throw new Error("chart XML part not found");
  return new TextDecoder().decode(chartPart.bytes);
}

export function mediaBytes(source: PptxSourceModel, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}

export function shapeWithoutTransform(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0].shapes.find(
    (node): node is SourceShape => node.kind === "shape" && node.transform === undefined,
  );
  if (shape === undefined) throw new Error("shape without transform not found");
  return shape;
}

export function findShapeByName(
  source: PptxSourceModel,
  name: string,
): SourceShapeNode | undefined {
  return source.slides
    .flatMap((slide) => slide.shapes)
    .find((shape) => shape.kind !== "raw" && shape.name === name);
}

export function findConnectorByName(source: PptxSourceModel, name: string): SourceConnector {
  const connector = source.slides
    .flatMap((slide) => slide.shapes)
    .find((shape): shape is SourceConnector => shape.kind === "connector" && shape.name === name);
  if (connector === undefined) throw new Error(`connector not found: ${name}`);
  return connector;
}

export function requireShape(shape: SourceShapeNode | undefined): SourceShape & {
  readonly transform: NonNullable<SourceShape["transform"]>;
} {
  if (!isShapeWithTransform(shape)) {
    throw new Error("transform shape not found");
  }
  return shape;
}

function isShapeWithTransform(
  shape: SourceShapeNode | undefined,
): shape is SourceShape & { readonly transform: NonNullable<SourceShape["transform"]> } {
  return shape?.kind === "shape" && shape.transform !== undefined;
}

export function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("handle not found");
  return handle;
}
