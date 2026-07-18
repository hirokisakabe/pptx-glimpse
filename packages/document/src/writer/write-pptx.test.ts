import { Buffer } from "node:buffer";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  addChart,
  addPicture,
  asEmu,
  asHundredthPt,
  asOoxmlAngle,
  asOoxmlPercent,
  asPartPath,
  asPt,
  asSourceNodeId,
  createComputedView,
  createPptx,
  readPptx,
  setSlideBackground,
  writePptx,
} from "../index.js";
import {
  addConnector,
  addEmptySlideFromLayout,
  addShape,
  addSlideNumber,
  addTable,
  addTextBox,
  clearParagraphProperties,
  clearTextRunProperties,
  deleteShape,
  deleteSlide,
  duplicateSlide,
  findParagraphBySourceHandle,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  moveSlide,
  replaceImageBytes,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setShapeFill,
  setShapeOutline,
  setTextRunProperties,
  type SourceConnector,
  type SourceImage,
  type SourceShape,
  type SourceShapeNode,
  type SourceTextRun,
  updateShapeTransform,
} from "../index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const RED_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=",
);
const BLUE_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
);
const GREEN_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEklEQVR4nGNk+M8AB0wIJj4OADLSAQcrNhPdAAAAAElFTkSuQmCC",
);

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function pngBytes(base64: string): Uint8Array {
  return new Uint8Array(Buffer.from(base64, "base64"));
}

function buildRoundTripFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst>` +
        `<p:sldId id="256" r:id="rIdSlide1"/>` +
        `<p:sldId id="257" r:id="rIdSlide2"/>` +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/slide2.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>` +
        `</Relationships>`,
    ),
    "ppt/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
    "docProps/custom.xml": xml(`<Properties><custom value="preserve-me"/></Properties>`),
  });
}

function buildMediaReplacementFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Replace Target"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `<Relationship Id="rIdImage2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>` +
        `</Relationships>`,
    ),
    "ppt/media/image1.png": RED_PNG,
    "ppt/media/image2.png": GREEN_PNG,
    "docProps/custom.xml": xml(`<Properties><custom value="preserve-me"/></Properties>`),
  });
}

function buildSlideTopologyFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` +
        `<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst>` +
        `<p:sldId id="256" r:id="rIdSlide1"/>` +
        `<p:sldId id="300" r:id="rIdSlide2"/>` +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Invisible Source"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slide One</a:t></a:r></a:p></p:txBody>` +
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `<p:timing><p:tnLst><p:par/></p:tnLst></p:timing>` +
        `</p:sld>`,
    ),
    "ppt/slides/slide2.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="20" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slide Two</a:t></a:r></a:p></p:txBody>` +
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>` +
        `<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/notesSlides/notesSlide1.xml": xml(
      `<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:notes>`,
    ),
    "ppt/notesSlides/_rels/notesSlide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/>` +
        `<Relationship Id="rIdNotesMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/comments/comment1.xml": xml(
      `<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
    ),
  });
}

function buildSlideTopologyFixtureWithRelationshipOverrides(): Uint8Array {
  const entries = unzipSync(buildSlideTopologyFixture());
  const relationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";
  const contentTypes = decoder
    .decode(entries["[Content_Types].xml"])
    .replace(`<Default Extension="rels" ContentType="${relationshipContentType}"/>`, "")
    .replace(
      `</Types>`,
      `<Override PartName="/_rels/.rels" ContentType="${relationshipContentType}"/>` +
        `<Override PartName="/ppt/_rels/presentation.xml.rels" ContentType="${relationshipContentType}"/>` +
        `<Override PartName="/ppt/slides/_rels/slide1.xml.rels" ContentType="${relationshipContentType}"/>` +
        `<Override PartName="/ppt/notesSlides/_rels/notesSlide1.xml.rels" ContentType="${relationshipContentType}"/>` +
        `</Types>`,
    );

  return zipSync({
    ...entries,
    "[Content_Types].xml": encoder.encode(contentTypes),
  });
}

function buildUnreferencedLayoutFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/><p:sldMasterId id="2147483649" r:id="rIdMaster2"/></p:sldMasterIdLst>` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<Relationship Id="rIdMaster2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">` +
        `<p:cSld name="Referenced"><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout2.xml": xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">` +
        `<p:cSld name="Unreferenced"><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout3.xml": xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="picTx">` +
        `<p:cSld name="Unused Master Layout"><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout3.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster1.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `<p:sldLayoutIdLst>` +
        `<p:sldLayoutId id="2147483649" r:id="rIdLayout1"/>` +
        `<p:sldLayoutId id="2147483650" r:id="rIdLayout2"/>` +
        `</p:sldLayoutIdLst>` +
        `</p:sldMaster>`,
    ),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `<Relationship Id="rIdLayout2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster2.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `<p:sldLayoutIdLst><p:sldLayoutId id="2147483651" r:id="rIdLayout3"/></p:sldLayoutIdLst>` +
        `</p:sldMaster>`,
    ),
    "ppt/slideMasters/_rels/slideMaster2.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdLayout3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml"/>` +
        `</Relationships>`,
    ),
  });
}

function buildLayoutShowRoundTripFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": xml(
      `<p:sldLayout show="0" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster1.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldMaster>`,
    ),
  });
}

function buildTextEditFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>` +
      `<a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1" i="1" sz="2400"><a:latin typeface="Aptos"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ea typeface="Noto Sans JP"/></a:rPr><a:t>Original</a:t></a:r><a:r><a:t xml:space="preserve"> Keep </a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`,
  );
}

function buildSlotHandleTextEditFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>` +
      `<p:sp><p:nvSpPr><p:cNvPr name="No Id Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original</a:t></a:r></a:p></p:txBody></p:sp>`,
  );
}

function buildMultipleTextEditFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>First paragraph</a:t></a:r></a:p>` +
      `<a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="11" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>Other shape</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`,
  );
}

function buildNumericLikeTextFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Numeric Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>007</a:t></a:r><a:r><a:t>1e5</a:t></a:r><a:r><a:t>12.50</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="11" name="Edit Me"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>Original</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`,
  );
}

function buildShapeStyleFixture(): Uint8Array {
  return buildTextEditFixtureFromSlide(
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Styled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>` +
      `<a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"/>` +
      `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="dash"/></a:ln>` +
      `</p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Styled</a:t></a:r></a:p></p:txBody>` +
      `</p:sp>` +
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="20" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
      `<p:spPr>` +
      `<a:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></a:xfrm>` +
      `<a:prstGeom prst="straightConnector1"/>` +
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:tailEnd type="triangle" w="med" len="med"/></a:ln>` +
      `</p:spPr>` +
      `</p:cxnSp>`,
  );
}

function buildTextEditFixtureFromSlide(slideSpTree: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>${slideSpTree}</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
    "docProps/custom.xml": xml(`<Properties><custom value="preserve-me"/></Properties>`),
    "ppt/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9, 8, 7]),
  });
}

function buildShapeDeleteFixture(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Delete Me"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Remove</a:t></a:r></a:p></p:txBody>` +
        `</p:sp>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Keep Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="30" name="Keep Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Keep</a:t></a:r></a:p></p:txBody>` +
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `<p:timing><p:tnLst><p:par/></p:tnLst></p:timing>` +
        `</p:sld>`,
    ),
    "docProps/custom.xml": xml(`<Properties><custom value="preserve-me"/></Properties>`),
  });
}

function getEntry(output: Uint8Array, path: string): Uint8Array {
  const entry = unzipSync(output)[path];
  if (entry === undefined) throw new Error(`missing zip entry: ${path}`);
  return entry;
}

function findTextRun(source: ReturnType<typeof readPptx>, text: string): SourceTextRun {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      for (const paragraph of shape.textBody?.paragraphs ?? []) {
        const run = paragraph.runs.find((candidate) => candidate.text === text);
        if (run !== undefined) return run;
      }
    }
  }
  throw new Error(`text run not found: ${text}`);
}

describe("writePptx - from-scratch builder", () => {
  it("preserves alternating shape and picture sibling order after write and reread", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const shape = (name: string) => ({
      geometry: { kind: "preset" as const, preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name,
    });

    source = addShape(source, slideHandle, shape("First shape"));
    source = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name: "Middle picture",
    });
    source = addShape(source, slideHandle, shape("Last shape"));

    const output = writePptx(source);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const spTreeXml = slideXml.slice(
      slideXml.indexOf("<p:spTree"),
      slideXml.indexOf("</p:spTree>"),
    );

    expect(spTreeXml.match(/<p:(?:sp|pic|graphicFrame)(?=[ >])/g)).toEqual([
      "<p:sp",
      "<p:pic",
      "<p:sp",
    ]);
    expect(readPptx(output).slides[0]?.shapes.map((item) => item.kind)).toEqual([
      "shape",
      "image",
      "shape",
    ]);

    const persisted = readPptx(output);
    const firstShapeHandle = persisted.slides[0]?.shapes[0]?.handle;
    if (firstShapeHandle === undefined) throw new Error("first shape handle should exist");
    const deletedOutput = writePptx(deleteShape(persisted, firstShapeHandle));
    const deletedSlideXml = decoder.decode(getEntry(deletedOutput, "ppt/slides/slide1.xml"));
    const deletedSpTreeXml = deletedSlideXml.slice(
      deletedSlideXml.indexOf("<p:spTree"),
      deletedSlideXml.indexOf("</p:spTree>"),
    );
    expect(deletedSpTreeXml.match(/<p:(?:sp|pic|graphicFrame)(?=[ >])/g)).toEqual([
      "<p:pic",
      "<p:sp",
    ]);
  });

  it("preserves shape, picture, chart, and table sibling order after write and reread", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name: "Ordered shape",
    });
    source = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name: "Ordered picture",
    });
    source = addChart(source, slideHandle, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: ["A"], values: [1] }],
      name: "Ordered chart",
    });
    source = addTable(source, slideHandle, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      columnWidths: [asEmu(1000)],
      rows: [{ height: asEmu(1000), cells: [{ text: "Cell" }] }],
      name: "Ordered table",
    });

    const output = writePptx(source);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const spTreeXml = slideXml.slice(
      slideXml.indexOf("<p:spTree"),
      slideXml.indexOf("</p:spTree>"),
    );

    expect(spTreeXml.match(/<p:(?:sp|pic|graphicFrame)(?=[ >])/g)).toEqual([
      "<p:sp",
      "<p:pic",
      "<p:graphicFrame",
      "<p:graphicFrame",
    ]);
    expect(readPptx(output).slides[0]?.shapes.map((item) => item.kind)).toEqual([
      "shape",
      "image",
      "chart",
      "table",
    ]);
  });

  it("writes all native chart types with editable workbooks and consistent package metadata", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const types = ["bar", "line", "pie", "area", "doughnut", "radar"] as const;
    const edited = types.reduce(
      (current, chartType, index) =>
        addChart(current, slideHandle, {
          chartType,
          offsetX: asEmu(index * 1200000),
          offsetY: asEmu(200000),
          width: asEmu(1100000),
          height: asEmu(1800000),
          title: `${chartType} title`,
          showLegend: true,
          legendPosition: "b",
          series: [
            { name: "Revenue", categories: ["Q1", "Q2"], values: [10, 20], color: "#4472C4" },
            { name: "Cost", categories: ["Q1", "Q2"], values: [7, 12], color: "ED7D31" },
          ],
          ...(chartType === "radar" ? { radarStyle: "filled" as const } : {}),
          ...(chartType === "doughnut" ? { holeSize: 60 } : {}),
          ...(chartType === "bar"
            ? {
                categoryAxis: { hidden: true, lineVisible: false, gridLinesVisible: false },
                valueAxis: { hidden: true, lineVisible: false, gridLinesVisible: false },
                plotLayout: { x: 0, y: 0, width: 1, height: 1 },
              }
            : {}),
        }),
      source,
    );

    const output = writePptx(edited);
    const archive = unzipSync(output);
    const contentTypes = decoder.decode(archive["[Content_Types].xml"]);
    const slideXml = decoder.decode(archive["ppt/slides/slide1.xml"]);
    const slideRels = decoder.decode(archive["ppt/slides/_rels/slide1.xml.rels"]);
    expect(contentTypes.match(/drawingml\.chart\+xml/g)).toHaveLength(6);
    expect(contentTypes.match(/spreadsheetml\.sheet/g)).toHaveLength(6);
    expect(slideXml.match(/<c:chart /g)).toHaveLength(6);
    expect(slideRels.match(/relationships\/chart/g)).toHaveLength(6);

    for (let index = 1; index <= 6; index += 1) {
      const chartXml = decoder.decode(archive[`ppt/charts/chart${index}.xml`]);
      const chartRels = decoder.decode(archive[`ppt/charts/_rels/chart${index}.xml.rels`]);
      const workbook = archive[`ppt/embeddings/Microsoft_Excel_Worksheet${index}.xlsx`];
      expect(chartXml).toContain(`<c:${types[index - 1]}Chart>`);
      expect(chartXml).toContain("Sheet1!$B$2:$B$3");
      expect(chartXml).toContain(`<c:v>Revenue</c:v>`);
      expect(chartXml).toContain(`<c:v>Q2</c:v>`);
      expect(chartXml).toContain(`<c:v>20</c:v>`);
      expect(chartRels).toContain(`Target="../embeddings/Microsoft_Excel_Worksheet${index}.xlsx"`);
      const worksheet = decoder.decode(unzipSync(workbook)["xl/worksheets/sheet1.xml"]);
      expect(worksheet).toContain(`<t>Revenue</t>`);
      expect(worksheet).toContain(`<t>Q2</t>`);
      expect(worksheet).toContain(`<c r="B3"><v>20</v></c>`);
    }
    expect(decoder.decode(archive["ppt/charts/chart1.xml"])).toContain(`<c:manualLayout>`);
    expect(decoder.decode(archive["ppt/charts/chart1.xml"])).toContain(`<c:delete val="1"/>`);
    expect(decoder.decode(archive["ppt/charts/chart5.xml"])).toContain(`<c:holeSize val="60"/>`);
    expect(decoder.decode(archive["ppt/charts/chart6.xml"])).toContain(
      `<c:radarStyle val="filled"/>`,
    );

    const reread = readPptx(output);
    expect(reread.diagnostics).toEqual([]);
    expect(createComputedView(reread).slides[0]?.elements.map((element) => element.kind)).toEqual(
      types.map(() => "chart"),
    );
    expect(
      createComputedView(reread).slides[0]?.elements.map((element) =>
        element.kind === "chart" ? element.chartData?.chartType : undefined,
      ),
    ).toEqual(types);
  });

  it("writes typed chart, axis, series, marker, and data point formatting", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const color = (hex: string, alpha?: number) => ({
      kind: "srgb" as const,
      hex,
      ...(alpha === undefined
        ? {}
        : { transforms: [{ kind: "alpha" as const, value: asOoxmlPercent(alpha) }] }),
    });
    const solid = (hex: string, alpha?: number) => ({
      kind: "solid" as const,
      color: color(hex, alpha),
    });

    const formatted = addChart(source, slideHandle, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(4000000),
      height: asEmu(2500000),
      title: "Formatted chart",
      titleStyle: {
        fontFace: "Aptos Display",
        fontSize: asPt(18),
        color: color("112233"),
        bold: true,
        italic: true,
      },
      displayBlanksAs: "span",
      roundedCorners: true,
      chartArea: {
        fill: solid("F0F0F0", 80000),
        outline: { width: asEmu(12700), fill: solid("223344"), dash: "dash" },
      },
      plotArea: {
        fill: solid("FFFFFF"),
        outline: { fill: { kind: "none" } },
      },
      categoryAxis: {
        hidden: true,
        majorTickMark: "inside",
        labelPosition: "low",
        numberFormat: { formatCode: "0.0%", sourceLinked: false },
        line: { width: asEmu(19050), fill: solid("445566") },
        majorGridline: { fill: solid("D9D9D9"), dash: "dot" },
        gridLinesVisible: true,
        textStyle: {
          fontFace: "Aptos",
          fontSize: asPt(9),
          color: color("667788"),
          bold: true,
          italic: false,
        },
        showMultiLevelLabels: false,
      },
      valueAxis: {
        hidden: true,
        lineVisible: false,
        majorTickMark: "outside",
        labelPosition: "none",
        numberFormat: { formatCode: "#,##0" },
        line: { fill: solid("FF0000") },
        gridLinesVisible: false,
      },
      plotLayout: { coordinateMode: "edge", x: 0, y: 0, width: 1, height: 1 },
      series: [
        {
          name: "Revenue",
          categories: ["Q1", "Q2"],
          values: [10, 20],
          fill: solid("4472C4"),
          outline: { width: asEmu(12700), fill: solid("203864"), dash: "solid" },
          dataPoints: [
            { index: 0, fill: solid("ED7D31"), outline: { fill: solid("843C0C") } },
            { index: 1, fill: solid("70AD47") },
          ],
        },
      ],
    });
    const withLine = addChart(formatted, slideHandle, {
      chartType: "line",
      offsetX: asEmu(4000000),
      offsetY: asEmu(0),
      width: asEmu(4000000),
      height: asEmu(2500000),
      series: [
        {
          categories: ["Q1", "Q2"],
          values: [4, 8],
          fill: { kind: "none" },
          outline: { width: asEmu(25400), fill: solid("5B9BD5"), dash: "dashDot" },
          marker: {
            symbol: "diamond",
            size: 9,
            fill: solid("FFC000"),
            outline: { fill: solid("7F6000") },
          },
        },
      ],
    });
    const withArea = addChart(withLine, slideHandle, {
      chartType: "area",
      offsetX: asEmu(0),
      offsetY: asEmu(2500000),
      width: asEmu(2500000),
      height: asEmu(2000000),
      series: [
        {
          categories: ["Q1", "Q2"],
          values: [2, 3],
          fill: solid("A5A5A5"),
          outline: { fill: solid("404040") },
        },
      ],
    });
    const withRadar = addChart(withArea, slideHandle, {
      chartType: "radar",
      offsetX: asEmu(2500000),
      offsetY: asEmu(2500000),
      width: asEmu(2500000),
      height: asEmu(2000000),
      series: [
        {
          categories: ["Q1", "Q2"],
          values: [5, 6],
          fill: solid("8064A2"),
          outline: { fill: solid("4F3B66") },
          marker: { symbol: "triangle", size: 7, fill: solid("8064A2") },
        },
      ],
    });
    const withPie = addChart(withRadar, slideHandle, {
      chartType: "pie",
      offsetX: asEmu(5000000),
      offsetY: asEmu(2500000),
      width: asEmu(1800000),
      height: asEmu(2000000),
      displayBlanksAs: "zero",
      series: [
        {
          categories: ["A", "B"],
          values: [1, 2],
          dataPoints: [
            { index: 0, fill: solid("C00000") },
            { index: 1, fill: solid("00B050"), outline: { fill: solid("006100") } },
          ],
        },
      ],
    });
    const edited = addChart(withPie, slideHandle, {
      chartType: "doughnut",
      offsetX: asEmu(6800000),
      offsetY: asEmu(2500000),
      width: asEmu(1800000),
      height: asEmu(2000000),
      displayBlanksAs: "gap",
      series: [
        {
          categories: ["A", "B"],
          values: [1, 2],
          dataPoints: [{ index: 0, fill: solid("00B0F0") }],
        },
      ],
    });

    const output = writePptx(edited);
    const archive = unzipSync(output);
    const barXml = decoder.decode(archive["ppt/charts/chart1.xml"]);
    const titleXml = /<c:title>.*?<\/c:title>/.exec(barXml)?.[0];
    const plotAreaXml = /<c:plotArea>.*?<\/c:plotArea>/.exec(barXml)?.[0];
    const barSeriesXml = /<c:ser>.*?<\/c:ser>/.exec(barXml)?.[0];
    expect(barXml).toContain(`<c:roundedCorners val="1"/>`);
    expect(barXml).toContain(`<c:dispBlanksAs val="span"/>`);
    expect(titleXml).toContain(`<a:rPr lang="en-US" sz="1800" b="1" i="1">`);
    expect(titleXml).toContain(`<a:srgbClr val="112233">`);
    expect(titleXml).toContain(`<a:latin typeface="Aptos Display"/>`);
    expect(plotAreaXml).toContain(
      `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`,
    );
    expect(barSeriesXml).toContain(
      `<c:spPr><a:solidFill><a:srgbClr val="4472C4"></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="203864"></a:srgbClr></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr>`,
    );
    expect(barXml).toContain(
      `<c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0"/><c:y val="0"/><c:w val="1"/><c:h val="1"/></c:manualLayout>`,
    );
    const categoryAxisXml = /<c:catAx>.*?<\/c:catAx>/.exec(barXml)?.[0];
    const valueAxisXml = /<c:valAx>.*?<\/c:valAx>/.exec(barXml)?.[0];
    expect(categoryAxisXml).toContain(`<c:delete val="1"/>`);
    expect(categoryAxisXml).toContain(`<c:majorTickMark val="in"/>`);
    expect(categoryAxisXml).toContain(`<c:tickLblPos val="low"/>`);
    expect(categoryAxisXml).toContain(`<c:numFmt formatCode="0.0%" sourceLinked="0"/>`);
    expect(categoryAxisXml).toContain(`<c:majorGridlines><c:spPr>`);
    expect(categoryAxisXml).toContain(`<a:ln w="19050">`);
    expect(categoryAxisXml).toContain(`<a:defRPr sz="900" b="1" i="0">`);
    expect(categoryAxisXml).toContain(`<c:noMultiLvlLbl val="1"/>`);
    expect(valueAxisXml).toContain(`<c:delete val="1"/>`);
    expect(valueAxisXml).toContain(`<c:majorTickMark val="out"/>`);
    expect(valueAxisXml).toContain(`<c:tickLblPos val="none"/>`);
    expect(valueAxisXml).toContain(`<c:numFmt formatCode="#,##0" sourceLinked="1"/>`);
    expect(valueAxisXml).not.toContain(`<c:majorGridlines>`);
    expect(valueAxisXml).toContain(`<a:ln><a:noFill/></a:ln>`);
    expect(barXml).toContain(
      `<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="843C0C"></a:srgbClr></a:solidFill></a:ln></c:spPr></c:dPt>`,
    );
    expect(barXml).toContain(
      `<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="70AD47"></a:srgbClr></a:solidFill></c:spPr></c:dPt>`,
    );
    expect(barXml).toContain(
      `</c:chart><c:spPr><a:solidFill><a:srgbClr val="F0F0F0"><a:alpha val="80000"/></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="223344"></a:srgbClr></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr><c:externalData`,
    );
    const lineXml = decoder.decode(archive["ppt/charts/chart2.xml"]);
    expect(lineXml).toContain(`<c:symbol val="diamond"/><c:size val="9"/>`);
    expect(lineXml).toContain(
      `<c:spPr><a:solidFill><a:srgbClr val="FFC000"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="7F6000"></a:srgbClr></a:solidFill></a:ln></c:spPr>`,
    );
    expect(lineXml).toContain(`<a:prstDash val="dashDot"/>`);
    const areaXml = decoder.decode(archive["ppt/charts/chart3.xml"]);
    expect(areaXml).toContain(
      `<a:solidFill><a:srgbClr val="A5A5A5"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="404040"></a:srgbClr></a:solidFill></a:ln>`,
    );
    const radarXml = decoder.decode(archive["ppt/charts/chart4.xml"]);
    expect(radarXml).toContain(`<c:symbol val="triangle"/><c:size val="7"/>`);
    expect(radarXml).toContain(
      `<a:solidFill><a:srgbClr val="8064A2"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="4F3B66"></a:srgbClr></a:solidFill></a:ln>`,
    );
    const pieXml = decoder.decode(archive["ppt/charts/chart5.xml"]);
    expect(pieXml).toContain(`<c:dispBlanksAs val="zero"/>`);
    expect(pieXml).toContain(
      `<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00B050"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="006100"></a:srgbClr></a:solidFill></a:ln></c:spPr></c:dPt>`,
    );
    const doughnutXml = decoder.decode(archive["ppt/charts/chart6.xml"]);
    expect(doughnutXml).toContain(`<c:dispBlanksAs val="gap"/>`);
    expect(doughnutXml).toContain(
      `<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="00B0F0"></a:srgbClr></a:solidFill></c:spPr></c:dPt>`,
    );

    const reread = readPptx(output);
    expect(reread.diagnostics).toEqual([]);
    expect(
      reread.packageGraph.parts.filter((part) =>
        /^ppt\/charts\/chart\d+\.xml$/.test(part.partPath),
      ),
    ).toHaveLength(6);
    expect(
      reread.packageGraph.parts.filter((part) => part.partPath.includes("/embeddings/")),
    ).toHaveLength(6);
    expect(
      createComputedView(reread).slides[0]?.elements.map((element) =>
        element.kind === "chart" ? element.chartData?.chartType : undefined,
      ),
    ).toEqual(["bar", "line", "area", "radar", "pie", "doughnut"]);
  });

  it("avoids the shape-tree root ID and rejects package-breaking chart inputs", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addChart(source, source.slides[0].handle!, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: [" A "], values: [1], name: " Series " }],
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const chartXml = decoder.decode(getEntry(output, "ppt/charts/chart1.xml"));
    const worksheet = decoder.decode(
      unzipSync(getEntry(output, "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"))[
        "xl/worksheets/sheet1.xml"
      ],
    );
    expect(slideXml).toContain(`<p:cNvPr id="31" name="Chart 31"/>`);
    expect(chartXml).toContain(`<c:v xml:space="preserve"> Series </c:v>`);
    expect(worksheet).toContain(`<t xml:space="preserve"> A </t>`);

    const valid = {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: ["A"], values: [1] }],
    } as const;
    expect(() =>
      addChart(source, source.slides[0].handle!, { ...valid, title: "bad\u0000" }),
    ).toThrow(/forbidden by XML 1.0/);
    const invalidFillKind = { ...valid, chartArea: { fill: { kind: "none" as const } } };
    Reflect.set(invalidFillKind.chartArea.fill, "kind", "gradient");
    expect(() => addChart(source, source.slides[0].handle!, invalidFillKind)).toThrow(
      /unsupported chartArea\.fill\.kind/,
    );
    const invalidColorKind = {
      ...valid,
      chartArea: {
        fill: { kind: "solid" as const, color: { kind: "srgb" as const, hex: "FFFFFF" } },
      },
    };
    Reflect.set(invalidColorKind.chartArea.fill.color, "kind", "scheme");
    expect(() => addChart(source, source.slides[0].handle!, invalidColorKind)).toThrow(
      /unsupported chartArea\.fill\.color\.kind/,
    );
    expect(() =>
      addChart(source, source.slides[0].handle!, {
        ...valid,
        chartArea: { outline: { width: asEmu(20_116_801) } },
      }),
    ).toThrow(/width must be a finite EMU value/);
    expect(() =>
      addChart(source, source.slides[0].handle!, {
        ...valid,
        chartArea: { outline: { width: asEmu(0.5) } },
      }),
    ).toThrow(/width must be a finite EMU value/);
    expect(() =>
      addChart(source, source.slides[0].handle!, {
        ...valid,
        title: "Too small",
        titleStyle: { fontSize: asPt(0.5) },
      }),
    ).toThrow(/fontSize must be from 1 through 4000 points/);
  });

  it("writes a new presentation after adding a text box through public APIs", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const edited = addTextBox(source, slideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      text: "Hello from scratch",
    });

    const reread = readPptx(writePptx(edited));
    expect(reread.diagnostics).toEqual([]);
    expect(reread.presentation.slidePartPaths).toEqual([asPartPath("ppt/slides/slide1.xml")]);
    expect(reread.slides).toHaveLength(1);
    expect(reread.slideLayouts).toHaveLength(1);
    expect(reread.slideMasters).toHaveLength(1);
    expect(reread.themes).toHaveLength(1);
    expect(findTextRun(reread, "Hello from scratch").text).toBe("Hello from scratch");

    const computed = createComputedView(reread);
    expect(computed.slideSize).toEqual({ width: asEmu(9144000), height: asEmu(5143500) });
    expect(computed.slides[0]?.elements).toHaveLength(1);
  });

  it("authors a named master and layout with backgrounds, objects, slide numbers, and margins", () => {
    let source = createPptx({
      slideMaster: { name: "Product Master", background: { kind: "image", bytes: BLUE_PNG } },
      slideLayout: {
        name: "Product Blank",
        margin: {
          left: asEmu(120000),
          right: asEmu(130000),
          top: asEmu(140000),
          bottom: asEmu(150000),
        },
      },
    });
    const masterHandle = source.slideMasters[0]?.handle;
    const layout = source.slideLayouts[0];
    if (masterHandle === undefined || layout?.handle === undefined) {
      throw new Error("createPptx should create master and layout handles");
    }
    source = addTextBox(source, masterHandle, {
      offsetX: asEmu(200000),
      offsetY: asEmu(100000),
      width: asEmu(1800000),
      height: asEmu(400000),
      text: "Master title",
    });
    source = addShape(source, masterHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(5000000),
      width: asEmu(9144000),
      height: asEmu(143500),
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
    source = addConnector(source, masterHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(200000),
      offsetY: asEmu(700000),
      width: asEmu(1800000),
      height: asEmu(1),
    });
    source = addPicture(source, masterHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(8200000),
      offsetY: asEmu(100000),
      width: asEmu(600000),
      height: asEmu(300000),
    });
    source = addSlideNumber(source, masterHandle, {
      offsetX: asEmu(8200000),
      offsetY: asEmu(4700000),
      width: asEmu(600000),
      height: asEmu(300000),
      align: "right",
      properties: { fontFace: "Aptos", fontSize: asPt(10), color: { kind: "srgb", hex: "334455" } },
    });
    source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    source = addTextBox(source, source.slides[1].handle!, {
      offsetX: asEmu(500000),
      offsetY: asEmu(500000),
      width: asEmu(2000000),
      height: asEmu(500000),
      text: "Uses layout margins",
      body: { marginRight: asEmu(999000) },
    });

    const output = writePptx(source);
    const masterXml = decoder.decode(getEntry(output, "ppt/slideMasters/slideMaster1.xml"));
    const masterRels = decoder.decode(
      getEntry(output, "ppt/slideMasters/_rels/slideMaster1.xml.rels"),
    );
    const slide2Xml = decoder.decode(getEntry(output, "ppt/slides/slide2.xml"));
    const reread = readPptx(output);

    expect(source.presentation.slidePartPaths).toHaveLength(3);
    expect(source.slideMasters[0]).toMatchObject({ name: "Product Master" });
    expect(source.slideLayouts[0]).toMatchObject({
      name: "Product Blank",
      defaultTextBodyProperties: {
        marginLeft: 120000,
        marginRight: 130000,
        marginTop: 140000,
        marginBottom: 150000,
      },
    });
    expect(masterXml).toContain(`<p:cSld name="Product Master"><p:bg>`);
    expect(masterXml).toContain(`<a:blip r:embed="rId3"/>`);
    expect(masterXml).toContain(`type="slidenum"`);
    expect(masterXml.match(/<p:cNvPr id="[1-5]"/g)).toHaveLength(5);
    expect(masterRels).toContain(`Id="rId3"`);
    expect(masterRels).toContain(`Target="../media/image1.png"`);
    expect(masterRels).toContain(`Id="rId4"`);
    expect(masterRels).toContain(`Target="../media/image2.png"`);
    expect(slide2Xml).toContain(`lIns="120000"`);
    expect(slide2Xml).toContain(`rIns="999000"`);
    expect(slide2Xml).toContain(`tIns="140000"`);
    expect(slide2Xml).toContain(`bIns="150000"`);
    expect(reread.diagnostics).toEqual([]);
    expect(reread.slideMasters[0]).toMatchObject({ name: "Product Master" });
    expect(reread.slideLayouts[0]).toMatchObject({ name: "Product Blank" });
    expect(reread.slideMasters[0]?.shapes).toHaveLength(5);
    expect(createComputedView(reread).slides.map((slide) => slide.elements.length)).toEqual([
      5, 6, 5,
    ]);
    expect(reread.packageGraph.contentTypes.defaults).toContainEqual({
      extension: "png",
      contentType: "image/png",
    });
  });

  it("authors a solid master background", () => {
    const output = writePptx(
      createPptx({
        slideMaster: {
          name: "Solid Master",
          background: { kind: "solid", color: { kind: "srgb", hex: "F8FAFC" } },
        },
      }),
    );
    const reread = readPptx(output);
    expect(reread.slideMasters[0]).toMatchObject({
      name: "Solid Master",
      background: {
        kind: "fill",
        fill: { kind: "solid", color: { kind: "srgb", hex: "F8FAFC" } },
      },
    });
  });

  it("authors solid, linear, radial, PNG, and JPEG backgrounds on individual slides", () => {
    let source = createPptx();
    const masterHandle = source.slideMasters[0]?.handle;
    const layout = source.slideLayouts[0];
    if (masterHandle === undefined || layout === undefined) {
      throw new Error("createPptx should create a master and layout");
    }
    source = addShape(source, masterHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100000),
      offsetY: asEmu(100000),
      width: asEmu(500000),
      height: asEmu(500000),
      fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
    });
    for (let index = 0; index < 4; index += 1) {
      source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    }
    const handles = source.slides.map((slide) => slide.handle);
    if (handles.some((handle) => handle === undefined)) {
      throw new Error("authored slides should have handles");
    }

    source = setSlideBackground(source, handles[0]!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    source = setSlideBackground(source, handles[1]!, {
      kind: "gradient",
      gradientType: "linear",
      angle: asOoxmlAngle(2700000),
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "0000FF" } },
      ],
    });
    source = setSlideBackground(source, handles[2]!, {
      kind: "gradient",
      gradientType: "radial",
      centerX: asOoxmlPercent(25000),
      centerY: asOoxmlPercent(75000),
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFFFF" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "00AA44" } },
      ],
    });
    source = setSlideBackground(source, handles[3]!, { kind: "image", bytes: BLUE_PNG });
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]);
    source = setSlideBackground(source, handles[4]!, { kind: "image", bytes: jpeg });

    const output = writePptx(source);
    const archive = unzipSync(output);
    const slideXml = Array.from({ length: 5 }, (_, index) =>
      decoder.decode(archive[`ppt/slides/slide${index + 1}.xml`]),
    );
    const pngRels = decoder.decode(archive["ppt/slides/_rels/slide4.xml.rels"]);
    const jpegRels = decoder.decode(archive["ppt/slides/_rels/slide5.xml.rels"]);
    const contentTypes = decoder.decode(archive["[Content_Types].xml"]);
    const reread = readPptx(output);

    expect(source.slides.map((slide) => slide.background?.kind)).toEqual([
      "fill",
      "fill",
      "fill",
      "fill",
      "fill",
    ]);
    expect(slideXml[0]).toContain(
      `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill>`,
    );
    expect(slideXml[1]).toContain(`<a:lin ang="2700000" scaled="1"/>`);
    expect(slideXml[2]).toContain(
      `<a:path path="circle"><a:fillToRect l="25000" t="75000" r="75000" b="25000"/></a:path>`,
    );
    for (const xml of slideXml) {
      expect(xml.indexOf("<p:bg>")).toBeLessThan(xml.indexOf("<p:spTree>"));
    }
    expect(pngRels).toContain(`Target="../media/image1.png"`);
    expect(jpegRels).toContain(`Target="../media/image1.jpeg"`);
    expect(archive["ppt/media/image1.png"]).toEqual(BLUE_PNG);
    expect(archive["ppt/media/image1.jpeg"]).toEqual(jpeg);
    expect(contentTypes).toContain(`<Default Extension="png" ContentType="image/png"/>`);
    expect(contentTypes).toContain(`<Default Extension="jpeg" ContentType="image/jpeg"/>`);
    expect(reread.diagnostics).toEqual([]);
    expect(reread.slides[1]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "gradient", gradientType: "linear", angle: 2700000 },
    });
    expect(reread.slides[2]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "gradient", gradientType: "radial", centerX: 0.25, centerY: 0.75 },
    });
    expect(reread.slides[3]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "image", blipRelationshipId: "rId2" },
    });
    expect(createComputedView(reread).slides.every((slide) => slide.elements.length === 1)).toBe(
      true,
    );
    expect(
      createComputedView(reread).slides.every(
        (slide) => slide.elements[0]?.sourceLayer === "master",
      ),
    ).toBe(true);
  });

  it("rejects invalid slide background inputs and non-slide handles", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    const masterHandle = source.slideMasters[0]?.handle;
    if (slideHandle === undefined || masterHandle === undefined) {
      throw new Error("createPptx should create slide and master handles");
    }
    expect(() =>
      setSlideBackground(source, slideHandle, {
        kind: "gradient",
        gradientType: "linear",
        angle: asOoxmlAngle(0),
        stops: [],
      }),
    ).toThrow(/at least two/);
    expect(() =>
      setSlideBackground(source, slideHandle, {
        kind: "image",
        bytes: new Uint8Array([1, 2, 3]),
      }),
    ).toThrow(/unsupported or unknown image format/);
    expect(() =>
      setSlideBackground(source, masterHandle, {
        kind: "solid",
        color: { kind: "srgb", hex: "FFFFFF" },
      }),
    ).toThrow(/slide handle was not found/);

    expect(() => {
      Reflect.apply(writePptx, undefined, [
        {
          ...source,
          edits: [
            {
              kind: "setSlideBackground",
              slidePartPath: source.slides[0].partPath,
              relationshipId: "rId2",
              xml: "<p:bg/>",
            },
          ],
        },
      ]);
    }).toThrow(/relationship, media part, and content type must be provided together/);
  });

  it("preserves a non-standard PresentationML prefix when authoring a slide background", () => {
    const archive = unzipSync(writePptx(createPptx()));
    const slidePartPath = "ppt/slides/slide1.xml";
    const slideXml = decoder
      .decode(archive[slidePartPath])
      .replaceAll("p:", "x:")
      .replace("xmlns:p=", "xmlns:x=");
    const source = readPptx(zipSync({ ...archive, [slidePartPath]: encoder.encode(slideXml) }));
    const edited = setSlideBackground(source, source.slides[0].handle!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    const writtenSlideXml = decoder.decode(unzipSync(writePptx(edited))[slidePartPath]);

    expect(writtenSlideXml).toContain("<x:bg><x:bgPr>");
    expect(writtenSlideXml).not.toContain("<p:bg");
  });

  it("preserves a namespace declared locally on an existing slide background", () => {
    const archive = unzipSync(writePptx(createPptx()));
    const slidePartPath = "ppt/slides/slide1.xml";
    const presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const existingBackground =
      `<x:bg xmlns:x="${presentationNamespace}"><x:bgPr>` +
      `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/>` +
      `</x:bgPr></x:bg>`;
    const slideXml = decoder
      .decode(archive[slidePartPath])
      .replace("<p:spTree>", `${existingBackground}<p:spTree>`);
    const source = readPptx(zipSync({ ...archive, [slidePartPath]: encoder.encode(slideXml) }));
    const edited = setSlideBackground(source, source.slides[0].handle!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    const output = writePptx(edited);
    const writtenSlideXml = decoder.decode(unzipSync(output)[slidePartPath]);
    const reread = readPptx(output);

    expect(writtenSlideXml).toContain(`<x:bg xmlns:x="${presentationNamespace}"><x:bgPr>`);
    expect(reread.diagnostics).toEqual([]);
    expect(reread.slides[0].background).toMatchObject({
      kind: "fill",
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
  });

  it("rejects authoring a background when the slide has no shape tree", () => {
    const archive = unzipSync(writePptx(createPptx()));
    const slidePartPath = "ppt/slides/slide1.xml";
    const slideXml = decoder
      .decode(archive[slidePartPath])
      .replace(/<p:spTree>[\s\S]*<\/p:spTree>/, "");
    const source = readPptx(zipSync({ ...archive, [slidePartPath]: encoder.encode(slideXml) }));
    const edited = setSlideBackground(source, source.slides[0].handle!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });

    expect(() => writePptx(edited)).toThrow(/has no p:spTree/);
  });

  it("authors every supported object on a layout with part-unique slide-number fields", () => {
    let source = createPptx();
    const masterHandle = source.slideMasters[0]?.handle;
    const layoutHandle = source.slideLayouts[0]?.handle;
    if (masterHandle === undefined || layoutHandle === undefined) {
      throw new Error("createPptx should create master and layout handles");
    }
    source = addTextBox(source, layoutHandle, {
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(300),
      height: asEmu(100),
      text: "Layout text",
    });
    source = addShape(source, layoutHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(20),
      offsetY: asEmu(30),
      width: asEmu(300),
      height: asEmu(100),
    });
    source = addConnector(source, layoutHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(30),
      offsetY: asEmu(40),
      width: asEmu(300),
      height: asEmu(1),
    });
    source = addPicture(source, layoutHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(40),
      offsetY: asEmu(50),
      width: asEmu(300),
      height: asEmu(100),
    });
    source = addSlideNumber(source, layoutHandle, {
      offsetX: asEmu(50),
      offsetY: asEmu(60),
      width: asEmu(300),
      height: asEmu(100),
    });
    for (let index = 0; index < 4; index += 1) {
      source = addSlideNumber(source, layoutHandle, {
        offsetX: asEmu(60 + index),
        offsetY: asEmu(70),
        width: asEmu(300),
        height: asEmu(100),
      });
    }
    source = addSlideNumber(source, masterHandle, {
      offsetX: asEmu(50),
      offsetY: asEmu(60),
      width: asEmu(300),
      height: asEmu(100),
    });

    const output = writePptx(source);
    const layoutXml = decoder.decode(getEntry(output, "ppt/slideLayouts/slideLayout1.xml"));
    const masterXml = decoder.decode(getEntry(output, "ppt/slideMasters/slideMaster1.xml"));
    const layoutRels = decoder.decode(
      getEntry(output, "ppt/slideLayouts/_rels/slideLayout1.xml.rels"),
    );
    const layoutFieldIds = [...layoutXml.matchAll(/<a:fld id="([^"]+)" type="slidenum"/g)].map(
      (match) => match[1],
    );
    const masterFieldId = masterXml.match(/<a:fld id="([^"]+)" type="slidenum"/)?.[1];
    const reread = readPptx(output);

    expect(layoutXml).toContain("<p:sp>");
    expect(layoutXml).toContain("<p:cxnSp>");
    expect(layoutXml).toContain("<p:pic>");
    expect(layoutRels).toContain('Target="../media/image1.png"');
    expect(layoutFieldIds).toHaveLength(5);
    expect(new Set(layoutFieldIds).size).toBe(5);
    expect(masterFieldId).toBeDefined();
    expect(layoutFieldIds).not.toContain(masterFieldId);
    expect(reread.slideLayouts[0]?.shapes).toHaveLength(9);
    expect(reread.slideMasters[0]?.shapes).toHaveLength(1);
    expect(reread.diagnostics).toEqual([]);
  });

  it("rejects invalid master and layout authoring options", () => {
    expect(() => {
      Reflect.apply(createPptx, undefined, [{ slideMaster: { background: { kind: "pattern" } } }]);
    }).toThrow(/background\.kind must be solid or image/);
    expect(() => createPptx({ slideMaster: { name: "bad\u0000name" } })).toThrow(
      /forbidden in an XML attribute/,
    );
    expect(() => createPptx({ slideLayout: { name: "bad\nname" } })).toThrow(
      /forbidden in an XML attribute/,
    );
  });

  it("writes run hyperlinks for text boxes and shapes with slide-local relationships", () => {
    const source = createPptx();
    const firstSlideHandle = source.slides[0]?.handle;
    const layoutPartPath = source.slideLayouts[0]?.partPath;
    if (firstSlideHandle === undefined || layoutPartPath === undefined) {
      throw new Error("createPptx should create a slide and layout");
    }

    const withTextBox = addTextBox(source, firstSlideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "plain" },
            { text: "first", hyperlink: "https://example.com/first?a=1&b=2" },
            { text: " middle", properties: { bold: true } },
            { text: " repeated", hyperlink: "https://example.com/first?a=1&b=2" },
            { text: " second", hyperlink: "http://example.com/second" },
          ],
        },
      ],
    });
    const withShape = addShape(withTextBox, firstSlideHandle, {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(2286000),
      width: asEmu(3657600),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "shape plain" },
            { text: " shape link", hyperlink: "https://example.com/first?a=1&b=2" },
          ],
        },
      ],
    });
    const withSecondSlide = addEmptySlideFromLayout(withShape, { layoutPartPath });
    const secondSlideHandle = withSecondSlide.slides[1]?.handle;
    if (secondSlideHandle === undefined) throw new Error("expected a second slide");
    const edited = addTextBox(withSecondSlide, secondSlideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "second slide" },
            { text: " linked", hyperlink: "https://example.com/slide-two" },
          ],
        },
      ],
    });

    const firstSlideRelationships = edited.packageGraph.relationships.find(
      (relationships) => relationships.sourcePartPath === "ppt/slides/slide1.xml",
    );
    const secondSlideRelationships = edited.packageGraph.relationships.find(
      (relationships) => relationships.sourcePartPath === "ppt/slides/slide2.xml",
    );
    expect(firstSlideRelationships?.relationships).toMatchObject([
      { id: "rId1", target: "../slideLayouts/slideLayout1.xml" },
      {
        id: "rId2",
        target: "https://example.com/first?a=1&b=2",
        targetMode: "External",
      },
      { id: "rId3", target: "http://example.com/second", targetMode: "External" },
      {
        id: "rId4",
        target: "https://example.com/first?a=1&b=2",
        targetMode: "External",
      },
    ]);
    expect(secondSlideRelationships?.relationships).toMatchObject([
      { id: "rId1", target: "../slideLayouts/slideLayout1.xml" },
      { id: "rId2", target: "https://example.com/slide-two", targetMode: "External" },
    ]);

    const output = writePptx(edited);
    const firstSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const secondSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide2.xml"));
    const firstSlideRelsXml = decoder.decode(getEntry(output, "ppt/slides/_rels/slide1.xml.rels"));
    const secondSlideRelsXml = decoder.decode(getEntry(output, "ppt/slides/_rels/slide2.xml.rels"));

    expect([...firstSlideXml.matchAll(/<a:hlinkClick r:id="rId2"\/>/g)]).toHaveLength(2);
    expect([...firstSlideXml.matchAll(/<a:hlinkClick r:id="rId3"\/>/g)]).toHaveLength(1);
    expect([...firstSlideXml.matchAll(/<a:hlinkClick r:id="rId4"\/>/g)]).toHaveLength(1);
    expect([...firstSlideXml.matchAll(/<a:hlinkClick/g)]).toHaveLength(4);
    expect([...secondSlideXml.matchAll(/<a:hlinkClick r:id="rId2"\/>/g)]).toHaveLength(1);
    expect(firstSlideRelsXml).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/first?a=1&amp;b=2" TargetMode="External"',
    );
    expect(firstSlideRelsXml).toContain(
      'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/second" TargetMode="External"',
    );
    expect(firstSlideRelsXml).toContain(
      'Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/first?a=1&amp;b=2" TargetMode="External"',
    );
    expect(secondSlideRelsXml).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/slide-two" TargetMode="External"',
    );
  });

  it("rejects non-HTTP run hyperlinks for text boxes and shapes", () => {
    const source = createPptx();
    const handle = source.slides[0]?.handle;
    if (handle === undefined) throw new Error("createPptx should create a first slide");
    const base = {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      paragraphs: [{ runs: [{ text: "link", hyperlink: "mailto:test@example.com" }] }],
    } as const;

    expect(() => addTextBox(source, handle, base)).toThrow(
      "addTextBox: paragraphs[0].runs[0].hyperlink must be an absolute HTTP(S) URL",
    );
    expect(() =>
      addShape(source, handle, { ...base, geometry: { kind: "preset", preset: "rect" } }),
    ).toThrow("addShape: paragraphs[0].runs[0].hyperlink must be an absolute HTTP(S) URL");
  });

  it("writes added PNG and JPEG pictures with media parts, content types, and slide rels", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const withPng = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(1828800),
      height: asEmu(1371600),
      rotation: asOoxmlAngle(600000),
      crop: {
        left: asOoxmlPercent(1000),
        top: asOoxmlPercent(2000),
        right: asOoxmlPercent(3000),
        bottom: asOoxmlPercent(4000),
      },
      name: "Product PNG",
    });
    const edited = addPicture(withPng, slideHandle, {
      bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]),
      offsetX: asEmu(3200400),
      offsetY: asEmu(457200),
      width: asEmu(1828800),
      height: asEmu(1371600),
      name: "Product JPEG",
    });

    const output = writePptx(edited);
    const contentTypesXml = decoder.decode(getEntry(output, "[Content_Types].xml"));
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const slideRelsXml = decoder.decode(getEntry(output, "ppt/slides/_rels/slide1.xml.rels"));
    const reread = readPptx(output);

    expect(getEntry(output, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(getEntry(output, "ppt/media/image1.jpeg")).toEqual(
      new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]),
    );
    expect(contentTypesXml).toContain(`<Default Extension="png" ContentType="image/png"/>`);
    expect(contentTypesXml).toContain(`<Default Extension="jpeg" ContentType="image/jpeg"/>`);
    expect(slideRelsXml).toContain(
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`,
    );
    expect(slideRelsXml).toContain(
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>`,
    );
    expect(slideXml).toContain(`<p:cNvPr id="1" name="Product PNG"/>`);
    expect(slideXml).toContain(`<a:blip r:embed="rId2"/>`);
    expect(slideXml).toContain(`<a:srcRect l="1000" t="2000" r="3000" b="4000"/>`);
    expect(slideXml).toContain(`<a:xfrm rot="600000">`);
    expect(slideXml).toContain(`<p:cNvPr id="2" name="Product JPEG"/>`);
    expect(slideXml).toContain(`<a:blip r:embed="rId3"/>`);
    expect(reread.packageGraph.media).toEqual([
      { partPath: "ppt/media/image1.png", contentType: "image/png", bytes: RED_PNG },
      {
        partPath: "ppt/media/image1.jpeg",
        contentType: "image/jpeg",
        bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]),
      },
    ]);
    expect(reread.slides[0]?.shapes).toMatchObject([
      {
        kind: "image",
        name: "Product PNG",
        blipRelationshipId: "rId2",
        crop: { left: 1000, top: 2000, right: 3000, bottom: 4000 },
      },
      {
        kind: "image",
        name: "Product JPEG",
        blipRelationshipId: "rId3",
      },
    ]);
  });

  it("writes and rereads shape and picture shadow effects", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      name: "Outer Shadow Shape",
      effects: {
        glow: {
          radius: asEmu(12700),
          color: { kind: "srgb", hex: "FFCC00" },
        },
        outerShadow: {
          blurRadius: asEmu(40000),
          distance: asEmu(20000),
          direction: asOoxmlAngle(5400000),
          color: {
            kind: "srgb",
            hex: "112233",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(40000) }],
          },
          alignment: "ctr",
          rotateWithShape: false,
        },
      },
    });
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "ellipse" },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      name: "Inner Shadow Shape",
      effects: {
        innerShadow: {
          blurRadius: asEmu(50000),
          distance: asEmu(30000),
          direction: asOoxmlAngle(10800000),
          color: { kind: "srgb", hex: "445566" },
        },
      },
    });
    source = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      name: "Shadow Picture",
      effects: {
        innerShadow: {
          blurRadius: asEmu(60000),
          distance: asEmu(35000),
          direction: asOoxmlAngle(9000000),
          color: { kind: "srgb", hex: "778899" },
        },
        outerShadow: {
          blurRadius: asEmu(70000),
          distance: asEmu(45000),
          direction: asOoxmlAngle(13500000),
          color: {
            kind: "srgb",
            hex: "000000",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(25000) }],
          },
          alignment: "br",
          rotateWithShape: true,
        },
      },
    });

    expect(findShapeByName(source, "Outer Shadow Shape").effects).toMatchObject({
      glow: { radius: 12700 },
      outerShadow: {
        blurRadius: 40000,
        distance: 20000,
        direction: 5400000,
        alignment: "ctr",
        rotateWithShape: false,
        color: { kind: "srgb", hex: "112233", transforms: [{ kind: "alpha", value: 40000 }] },
      },
    });
    expect(findShapeByName(source, "Inner Shadow Shape").effects?.innerShadow).toMatchObject({
      blurRadius: 50000,
      distance: 30000,
      direction: 10800000,
      color: { kind: "srgb", hex: "445566" },
    });
    expect(findImageByName(source, "Shadow Picture")).toMatchObject({
      effects: {
        innerShadow: { blurRadius: 60000, distance: 35000, direction: 9000000 },
        outerShadow: {
          blurRadius: 70000,
          distance: 45000,
          direction: 13500000,
          alignment: "br",
          rotateWithShape: true,
        },
      },
    });

    const output = writePptx(source);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(reread.diagnostics).toEqual([]);
    expect(findImageByName(reread, "Shadow Picture")).toMatchObject(
      findImageByName(source, "Shadow Picture"),
    );
    expect(slideXml).toContain(
      `<a:effectLst><a:glow rad="12700"><a:srgbClr val="FFCC00"/></a:glow><a:outerShdw blurRad="40000" dist="20000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="112233"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
    );
    expect(slideXml).toContain(
      `<a:innerShdw blurRad="50000" dist="30000" dir="10800000"><a:srgbClr val="445566"/></a:innerShdw>`,
    );
    expect(slideXml).toContain(
      `<a:effectLst><a:innerShdw blurRad="60000" dist="35000" dir="9000000"><a:srgbClr val="778899"/></a:innerShdw><a:outerShdw blurRad="70000" dist="45000" dir="13500000" algn="br" rotWithShape="1"><a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
    );
  });

  it("keeps added pictures at the serialized shape-tree end and adds missing relationship namespace", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addPicture(source, source.slides[0].handle!, {
      bytes: RED_PNG,
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      name: "Appended Picture",
    });
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(slideXml).toContain(
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`,
    );
    expect(slideXml.indexOf(`name="Appended Picture"`)).toBeGreaterThan(
      slideXml.indexOf(`name="Keep Shape"`),
    );
  });

  it("writes formatted text box rPr, pPr, bodyPr, and xfrm from public APIs", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const edited = addTextBox(source, slideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(4572000),
      height: asEmu(1828800),
      rotation: asOoxmlAngle(5400000),
      name: "Formatted TextBox",
      body: {
        anchor: "middle",
        marginLeft: asEmu(91440),
        marginRight: asEmu(182880),
        marginTop: asEmu(45720),
        marginBottom: asEmu(68580),
        autoFit: "shape",
      },
      paragraphs: [
        {
          properties: {
            align: "center",
            marginLeft: asEmu(342900),
            indent: asEmu(-285750),
            lineSpacing: { type: "points", value: asHundredthPt(1800) },
            bullet: {
              type: "character",
              character: "•",
              fontFace: "Aptos",
              size: asOoxmlPercent(125000),
            },
          },
          runs: [
            {
              text: "Solid styled",
              properties: {
                fontFace: "Aptos",
                fontSize: asPt(28),
                color: { kind: "srgb", hex: "112233" },
                bold: true,
                italic: true,
                underline: { style: "dbl", color: { kind: "srgb", hex: "445566" } },
                strike: true,
                highlight: { kind: "srgb", hex: "ffff00" },
                glow: { radius: asEmu(25400), color: { kind: "srgb", hex: "00aaff" } },
                outline: { width: asEmu(12700), color: { kind: "srgb", hex: "aa00aa" } },
                charSpacing: 120,
              },
            },
            {
              text: " gradient",
              properties: {
                gradientFill: {
                  gradientType: "linear",
                  angle: asOoxmlAngle(2700000),
                  stops: [
                    { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "ff0000" } },
                    {
                      position: asOoxmlPercent(100000),
                      color: { kind: "srgb", hex: "0000ff" },
                    },
                  ],
                },
                baseline: "superscript",
              },
            },
          ],
        },
        {
          properties: {
            align: "right",
            lineSpacing: { type: "percent", value: asOoxmlPercent(90000) },
            bullet: {
              type: "auto-number",
              scheme: "alphaLcParenR",
              startAt: 3,
              fontFace: "Aptos",
              size: asOoxmlPercent(100000),
            },
          },
          runs: [
            {
              text: "Subscript line",
              properties: {
                baseline: { type: "percent", value: asOoxmlPercent(-12500) },
                underline: true,
              },
            },
          ],
        },
      ],
    });
    const withShape = addShape(edited, edited.slides[0].handle!, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(2743200),
      width: asEmu(4572000),
      height: asEmu(914400),
      name: "Auto-fit Shape",
      body: { autoFit: "shape" },
      paragraphs: [
        {
          properties: { bullet: { type: "none" } },
          runs: [{ text: "Shape text" }],
        },
      ],
    });
    const output = writePptx(withShape);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const added = requireShape(findShapeByName(reread, "Formatted TextBox"));
    const addedShape = requireShape(findShapeByName(reread, "Auto-fit Shape"));

    expect(reread.diagnostics).toEqual([]);
    expect(added.transform).toMatchObject({ rotation: 5400000 });
    expect(added.textBody?.properties).toMatchObject({
      anchor: "middle",
      marginLeft: 91440,
      marginRight: 182880,
      marginTop: 45720,
      marginBottom: 68580,
      autoFit: "spAutofit",
    });
    expect(added.textBody?.paragraphs).toHaveLength(2);
    expect(added.textBody?.paragraphs[0]?.properties).toMatchObject({
      align: "center",
      lineSpacing: { type: "pts", value: 1800 },
      marginLeft: 342900,
      indent: -285750,
      bullet: { type: "char", char: "•" },
      bulletFont: "Aptos",
      bulletSizePct: 125000,
    });
    expect(added.textBody?.paragraphs[0]?.runs.map((run) => run.text)).toEqual([
      "Solid styled",
      " gradient",
    ]);
    expect(added.textBody?.paragraphs[1]?.properties).toMatchObject({
      align: "right",
      lineSpacing: { type: "pct", value: 90000 },
      bullet: { type: "autoNum", scheme: "alphaLcParenR", startAt: 3 },
      bulletFont: "Aptos",
      bulletSizePct: 100000,
    });
    expect(added.textBody?.paragraphs[1]?.runs[0]?.properties?.baseline).toBe(-12.5);
    expect(addedShape.textBody?.properties?.autoFit).toBe("spAutofit");
    expect(addedShape.textBody?.paragraphs[0]?.properties?.bullet).toEqual({ type: "none" });
    expect(slideXml).toContain(`<a:xfrm rot="5400000">`);
    expect(slideXml).toContain(
      `<a:bodyPr wrap="square" anchor="ctr" lIns="91440" rIns="182880" tIns="45720" bIns="68580"><a:spAutoFit/></a:bodyPr>`,
    );
    expect(slideXml).toContain(
      `<a:pPr algn="ctr" marL="342900" indent="-285750"><a:lnSpc><a:spcPts val="1800"/></a:lnSpc><a:buSzPct val="125000"/><a:buFont typeface="Aptos"/><a:buChar char="•"/>`,
    );
    expect(slideXml).toContain(
      `<a:pPr algn="r"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buSzPct val="100000"/><a:buFont typeface="Aptos"/><a:buAutoNum type="alphaLcParenR" startAt="3"/>`,
    );
    expect(slideXml).toContain(`<a:pPr><a:buNone/></a:pPr>`);
    expect(slideXml).toContain(`b="1"`);
    expect(slideXml).toContain(`i="1"`);
    expect(slideXml).toContain(`u="dbl"`);
    expect(slideXml).toContain(`strike="sngStrike"`);
    expect(slideXml).toContain(`sz="2800"`);
    expect(slideXml).toContain(`spc="120"`);
    expect(slideXml).toContain(`<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`);
    expect(slideXml).toContain(
      `<a:uFill><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:uFill>`,
    );
    expect(slideXml).toContain(`<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>`);
    expect(slideXml).toContain(
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln>`,
    );
    expect(slideXml).toContain(
      `<a:effectLst><a:glow rad="25400"><a:srgbClr val="00AAFF"/></a:glow></a:effectLst>`,
    );
    expect(slideXml).toContain(`<a:latin typeface="Aptos"/>`);
    expect(slideXml).toContain(`<a:ea typeface="Aptos"/>`);
    expect(slideXml).toContain(`<a:cs typeface="Aptos"/>`);
    expect(slideXml).toContain(`<a:gradFill><a:gsLst>`);
    expect(slideXml).toContain(`<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>`);
    expect(slideXml).toContain(`<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>`);
    expect(slideXml).toContain(`<a:lin ang="2700000" scaled="1"/>`);
    expect(slideXml).toContain(`baseline="30000"`);
    expect(slideXml).toContain(`baseline="-12500"`);
  });

  it("writes alpha colors and linear or radial gradients consistently with the source model", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const withText = addTextBox(source, slideHandle, {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      name: "Alpha Text",
      paragraphs: [
        {
          runs: [
            {
              text: "alpha",
              properties: {
                color: {
                  kind: "srgb",
                  hex: "112233",
                  transforms: [{ kind: "alpha", value: asOoxmlPercent(0) }],
                },
                glow: {
                  radius: asEmu(12700),
                  color: {
                    kind: "srgb",
                    hex: "445566",
                    transforms: [{ kind: "alpha", value: asOoxmlPercent(100000) }],
                  },
                },
              },
            },
            {
              text: " gradient",
              properties: {
                gradientFill: {
                  gradientType: "linear",
                  angle: asOoxmlAngle(1800000),
                  stops: [
                    {
                      position: asOoxmlPercent(0),
                      color: {
                        kind: "srgb",
                        hex: "FF0000",
                        transforms: [{ kind: "alpha", value: asOoxmlPercent(25000) }],
                      },
                    },
                    {
                      position: asOoxmlPercent(100000),
                      color: { kind: "srgb", hex: "0000FF" },
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const withSolidAndLinearOutline = addShape(withText, withText.slides[0].handle!, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      name: "Alpha Solid",
      fill: {
        kind: "solid",
        color: {
          kind: "srgb",
          hex: "778899",
          transforms: [{ kind: "alpha", value: asOoxmlPercent(50000) }],
        },
      },
      outline: {
        fill: {
          kind: "gradient",
          gradientType: "linear",
          stops: [
            {
              position: asOoxmlPercent(0),
              color: {
                kind: "srgb",
                hex: "00AA00",
                transforms: [{ kind: "alpha", value: asOoxmlPercent(40000) }],
              },
            },
            {
              position: asOoxmlPercent(100000),
              color: { kind: "srgb", hex: "AA0000" },
            },
          ],
        },
      },
      effects: {
        glow: {
          radius: asEmu(25400),
          color: {
            kind: "srgb",
            hex: "ABCDEF",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(60000) }],
          },
        },
      },
    });
    const edited = addShape(
      withSolidAndLinearOutline,
      withSolidAndLinearOutline.slides[0].handle!,
      {
        geometry: { kind: "preset", preset: "ellipse" },
        offsetX: asEmu(900),
        offsetY: asEmu(1000),
        width: asEmu(1100),
        height: asEmu(1200),
        name: "Radial Shape",
        fill: {
          kind: "gradient",
          gradientType: "radial",
          centerX: asOoxmlPercent(25000),
          centerY: asOoxmlPercent(75000),
          stops: [
            { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFFFF" } },
            {
              position: asOoxmlPercent(100000),
              color: {
                kind: "srgb",
                hex: "000000",
                transforms: [{ kind: "alpha", value: asOoxmlPercent(75000) }],
              },
            },
          ],
        },
        outline: {
          fill: {
            kind: "gradient",
            gradientType: "radial",
            centerX: asOoxmlPercent(50000),
            centerY: asOoxmlPercent(50000),
            stops: [
              { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFF00" } },
              {
                position: asOoxmlPercent(100000),
                color: {
                  kind: "srgb",
                  hex: "00FFFF",
                  transforms: [{ kind: "alpha", value: asOoxmlPercent(30000) }],
                },
              },
            ],
          },
        },
      },
    );

    const authoredText = requireShape(findShapeByName(edited, "Alpha Text"));
    const authoredSolid = requireShape(findShapeByName(edited, "Alpha Solid"));
    const authoredRadial = requireShape(findShapeByName(edited, "Radial Shape"));
    const output = writePptx(edited);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(reread.diagnostics).toEqual([]);
    expect(authoredText.textBody?.paragraphs[0]?.runs[0]?.properties?.color).toEqual({
      kind: "srgb",
      hex: "112233",
      transforms: [{ kind: "alpha", value: 0 }],
    });
    expect(authoredSolid.fill).toEqual({
      kind: "solid",
      color: {
        kind: "srgb",
        hex: "778899",
        transforms: [{ kind: "alpha", value: 50000 }],
      },
    });
    expect(authoredSolid.effects?.glow?.color).toEqual({
      kind: "srgb",
      hex: "ABCDEF",
      transforms: [{ kind: "alpha", value: 60000 }],
    });
    expect(authoredRadial.fill).toMatchObject({
      kind: "gradient",
      gradientType: "radial",
      centerX: 0.25,
      centerY: 0.75,
    });
    expect(
      requireShape(findShapeByName(reread, "Alpha Text")).textBody?.paragraphs[0]?.runs[0]
        ?.properties?.color,
    ).toEqual(authoredText.textBody?.paragraphs[0]?.runs[0]?.properties?.color);
    expect(requireShape(findShapeByName(reread, "Alpha Solid")).fill).toEqual(authoredSolid.fill);
    expect(requireShape(findShapeByName(reread, "Alpha Solid")).outline).toEqual(
      authoredSolid.outline,
    );
    expect(requireShape(findShapeByName(reread, "Radial Shape")).fill).toEqual(authoredRadial.fill);
    expect(requireShape(findShapeByName(reread, "Radial Shape")).outline).toEqual(
      authoredRadial.outline,
    );
    expect(slideXml).toContain(
      `<a:solidFill><a:srgbClr val="112233"><a:alpha val="0"/></a:srgbClr></a:solidFill>`,
    );
    expect(slideXml).toContain(
      `<a:glow rad="12700"><a:srgbClr val="445566"><a:alpha val="100000"/></a:srgbClr></a:glow>`,
    );
    expect(slideXml).toContain(
      `<a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="25000"/></a:srgbClr></a:gs>`,
    );
    expect(slideXml).toContain(
      `<a:solidFill><a:srgbClr val="778899"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
    );
    expect(slideXml).toContain(
      `<a:glow rad="25400"><a:srgbClr val="ABCDEF"><a:alpha val="60000"/></a:srgbClr></a:glow>`,
    );
    expect(slideXml).toContain(
      `<a:path path="circle"><a:fillToRect l="25000" t="75000" r="75000" b="25000"/></a:path>`,
    );
    expect(slideXml).toContain(
      `<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`,
    );
    expect(slideXml.match(/<a:alpha val=/g)).toHaveLength(8);
  });

  it("writes preset geometry shapes with fill, line, glow, rotation, and text", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const withSolid = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1000),
      offsetY: asEmu(2000),
      width: asEmu(3000),
      height: asEmu(4000),
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
      name: "Solid Rect",
    });
    const edited = addShape(withSolid, withSolid.slides[0].handle!, {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      rotation: asOoxmlAngle(5400000),
      fill: {
        kind: "gradient",
        gradientType: "linear",
        angle: asOoxmlAngle(2700000),
        stops: [
          { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "ff0000" } },
          { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "0000ff" } },
        ],
      },
      outline: {
        width: asEmu(12700),
        fill: { kind: "solid", color: { kind: "srgb", hex: "00aa44" } },
        dash: "dash",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
      effects: {
        glow: { radius: asEmu(25400), color: { kind: "srgb", hex: "aa00aa" } },
      },
      body: { anchor: "middle" },
      paragraphs: [
        {
          properties: { align: "center" },
          runs: [{ text: "Shape label", properties: { bold: true } }],
        },
      ],
      name: "Styled Shape",
    });
    const lineShape = addShape(edited, edited.slides[0].handle!, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
      name: "Line Shape",
    });
    const ellipseShape = addShape(lineShape, lineShape.slides[0].handle!, {
      geometry: { kind: "preset", preset: "ellipse" },
      offsetX: asEmu(5),
      offsetY: asEmu(6),
      width: asEmu(7),
      height: asEmu(8),
      name: "Ellipse Shape",
    });
    const output = writePptx(ellipseShape);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const solid = requireShape(findShapeByName(reread, "Solid Rect"));
    const styled = requireShape(findShapeByName(reread, "Styled Shape"));

    expect(reread.diagnostics).toEqual([]);
    expect(solid).toMatchObject({
      geometry: { preset: "rect" },
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
    expect(styled).toMatchObject({
      geometry: { preset: "roundRect" },
      transform: { rotation: 5400000 },
      fill: {
        kind: "gradient",
        gradientType: "linear",
        angle: 2700000,
      },
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { kind: "srgb", hex: "00AA44" } },
        dashStyle: "dash",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
      effects: {
        glow: { radius: 25400, color: { kind: "srgb", hex: "AA00AA" } },
      },
    });
    expect(styled.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Shape label");
    expect(findShapeByName(reread, "Line Shape").geometry).toEqual({ preset: "line" });
    expect(findShapeByName(reread, "Ellipse Shape").geometry).toEqual({ preset: "ellipse" });
    expect(slideXml).toContain(`<a:prstGeom prst="rect"`);
    expect(slideXml).toContain(`<a:prstGeom prst="roundRect"`);
    expect(slideXml).toContain(`<a:prstGeom prst="line"`);
    expect(slideXml).toContain(`<a:prstGeom prst="ellipse"`);
    expect(slideXml).toContain(`<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`);
    expect(slideXml).toContain(`<a:gradFill><a:gsLst>`);
    expect(slideXml).toContain(`<a:lin ang="2700000" scaled="1"/>`);
    expect(slideXml).toContain(
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="00AA44"/></a:solidFill><a:prstDash val="dash"/>`,
    );
    expect(slideXml).toContain(`<a:headEnd type="oval" w="sm" len="sm"`);
    expect(slideXml).toContain(`<a:tailEnd type="triangle" w="med" len="lg"`);
    expect(slideXml).toContain(
      `<a:effectLst><a:glow rad="25400"><a:srgbClr val="AA00AA"/></a:glow></a:effectLst>`,
    );
    expect(slideXml).toContain(`<p:txBody>`);
    expect(slideXml).toContain(`<a:t>Shape label</a:t>`);
    expect(slideXml).toContain(`<a:xfrm rot="5400000">`);
  });

  it("writes and rereads adjusted, custom, flipped, and zero-extent shape geometry", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "roundRect", adjustValues: { adj: 25000 } },
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      flipHorizontal: true,
      name: "Adjusted Geometry",
    });
    source = addShape(source, slideHandle, {
      geometry: {
        kind: "custom",
        paths: [
          {
            width: 100,
            height: 100,
            commands: [
              { kind: "moveTo", x: 0, y: 100 },
              { kind: "lineTo", x: 50, y: 0 },
              { kind: "lineTo", x: 100, y: 100 },
              { kind: "close" },
            ],
          },
        ],
      },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      flipVertical: true,
      name: "Custom Geometry",
    });
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(0),
      height: asEmu(1200),
      flipHorizontal: true,
      flipVertical: true,
      name: "Zero Extent Line",
    });

    const output = writePptx(source);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(requireShape(findShapeByName(reread, "Adjusted Geometry"))).toMatchObject({
      geometry: { preset: "roundRect", adjustValues: { adj: 25000 } },
      transform: { flipHorizontal: true },
    });
    expect(requireShape(findShapeByName(reread, "Custom Geometry"))).toMatchObject({
      geometry: {
        kind: "custom",
        paths: [{ width: 100, height: 100, commands: "M 0 100 L 50 0 L 100 100 Z" }],
      },
      transform: { flipVertical: true },
    });
    expect(requireShape(findShapeByName(reread, "Zero Extent Line"))).toMatchObject({
      geometry: { preset: "line" },
      transform: { width: 0, height: 1200, flipHorizontal: true, flipVertical: true },
    });
    expect(slideXml).toContain(`<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>`);
    expect(slideXml).toContain(`<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>`);
    expect(slideXml).toContain(`<a:moveTo><a:pt x="0" y="100"/></a:moveTo>`);
    expect(slideXml).toContain(`<a:lnTo><a:pt x="50" y="0"/></a:lnTo>`);
    expect(slideXml).toContain(`<a:close/>`);
    expect(slideXml).toContain(`<a:xfrm flipH="1" flipV="1">`);
    expect(slideXml).toContain(`<a:ext cx="0" cy="1200"/>`);
  });

  it("writes custom slide size without fixed 16:9 metadata", () => {
    const source = createPptx({
      slideSize: { width: asEmu(7315200), height: asEmu(5486400) },
    });
    const output = writePptx(source);
    const reread = readPptx(output);

    expect(reread.presentation.slideSize).toEqual({
      width: asEmu(7315200),
      height: asEmu(5486400),
    });
    expect(decoder.decode(getEntry(output, "ppt/presentation.xml"))).toContain(
      `<p:sldSz cx="7315200" cy="5486400"/>`,
    );
    expect(decoder.decode(getEntry(output, "docProps/app.xml"))).not.toContain(
      "On-screen Show (16:9)",
    );
  });

  it("rejects invalid custom slide sizes", () => {
    expect(() =>
      createPptx({
        slideSize: { width: asEmu(Number.NaN), height: asEmu(5486400) },
      }),
    ).toThrow(/slideSize\.width/);
    expect(() =>
      createPptx({
        slideSize: { width: asEmu(7315200), height: asEmu(0) },
      }),
    ).toThrow(/slideSize\.height/);
  });
});

describe("writePptx - no-edit round-trip", () => {
  it("You can write no-edit PPTX from PptxSourceModel source and reload it.", () => {
    const input = buildRoundTripFixture();
    const original = readPptx(input);

    const output = writePptx(original);
    const reread = readPptx(output);

    expect(reread.presentation.slidePartPaths).toEqual(original.presentation.slidePartPaths);
    expect(reread.presentation.slideSize).toEqual(original.presentation.slideSize);
    expect(reread.slides.map((slide) => slide.partPath)).toEqual(
      original.slides.map((slide) => slide.partPath),
    );
  });

  it("Preserving relationship IDs and content types structurally", () => {
    const original = readPptx(buildRoundTripFixture());
    const reread = readPptx(writePptx(original));

    expect(reread.packageGraph.contentTypes).toEqual(original.packageGraph.contentTypes);
    expect(reread.packageGraph.relationships).toEqual(original.packageGraph.relationships);
  });

  it("Preserves p:sldLayout@show structurally in no-edit round-trip", () => {
    const input = buildLayoutShowRoundTripFixture();
    const source = readPptx(input);
    const output = writePptx(source);
    const reread = readPptx(output);

    expect(source.slideLayouts[0]?.show).toBe(false);
    expect(decoder.decode(getEntry(output, "ppt/slideLayouts/slideLayout1.xml"))).toContain(
      `show="0"`,
    );
    expect(reread.slideLayouts[0]?.show).toBe(false);
  });

  it("Preserving media bytes and unsupported raw package material", () => {
    const input = buildRoundTripFixture();
    const source = readPptx(input);
    const output = writePptx(source);

    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toBe(
      decoder.decode(getEntry(input, "docProps/custom.xml")),
    );
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).toBe(
      decoder.decode(getEntry(input, "ppt/slides/slide1.xml")),
    );
  });

  it("Replaces one image media part while preserving relationships, content types, and other media bytes", () => {
    const input = buildMediaReplacementFixture();
    const source = readPptx(input);
    const image = source.slides[0]?.shapes.find((shape) => shape.kind === "image");
    if (image === undefined) throw new Error("image fixture was not parsed");
    const edited = replaceImageBytes(source, image.handle!, BLUE_PNG);
    const output = writePptx(edited);
    const reread = readPptx(output);

    expect(getEntry(output, "ppt/media/image1.png")).toEqual(BLUE_PNG);
    expect(getEntry(output, "ppt/media/image2.png")).toEqual(GREEN_PNG);
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toBe(
      decoder.decode(getEntry(input, "docProps/custom.xml")),
    );
    expect(reread.packageGraph.contentTypes).toEqual(source.packageGraph.contentTypes);
    expect(reread.packageGraph.relationships).toEqual(source.packageGraph.relationships);
    expect(
      reread.packageGraph.media.find((part) => part.partPath === "ppt/media/image1.png"),
    ).toMatchObject({
      contentType: "image/png",
      bytes: BLUE_PNG,
    });
    expect(
      reread.packageGraph.media.find((part) => part.partPath === "ppt/media/image2.png"),
    ).toMatchObject({
      contentType: "image/png",
      bytes: GREEN_PNG,
    });
  });

  it("Can write xml raw package material in serializable range", () => {
    const source = readPptx(buildRoundTripFixture());
    const withXmlRaw = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        contentTypes: {
          ...source.packageGraph.contentTypes,
          overrides: [
            ...source.packageGraph.contentTypes.overrides,
            { partName: "customXml/item1.xml", contentType: "application/xml" },
          ],
        },
        parts: [
          ...source.packageGraph.parts,
          { partPath: "customXml/item1.xml", contentType: "application/xml" },
        ],
        rawParts: [
          ...(source.packageGraph.rawParts ?? []),
          {
            kind: "xml",
            partPath: "customXml/item1.xml",
            contentType: "application/xml",
            xml: {
              name: "x:root",
              attributes: { "xmlns:x": "urn:test", "a:flag": "A&B" },
              children: [{ name: "x:child", text: "nested < text" }],
            },
          },
        ],
      },
    };

    expect(decoder.decode(getEntry(writePptx(withXmlRaw), "customXml/item1.xml"))).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<x:root xmlns:x="urn:test" a:flag="A&amp;B"><x:child>nested &lt; text</x:child></x:root>`,
    );
  });

  it("Reject mixed content of xml raw package material because order cannot be maintained", () => {
    const source = readPptx(buildRoundTripFixture());
    const withMixedContentRaw = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        parts: [
          ...source.packageGraph.parts,
          { partPath: "customXml/item1.xml", contentType: "application/xml" },
        ],
        rawParts: [
          ...(source.packageGraph.rawParts ?? []),
          {
            kind: "xml",
            partPath: "customXml/item1.xml",
            contentType: "application/xml",
            xml: {
              name: "x:root",
              text: "pre",
              children: [{ name: "x:child" }],
            },
          },
        ],
      },
    };

    expect(() => writePptx(withMixedContentRaw)).toThrow(/mixed text\/element content/);
  });

  it("Verify structural preservation rather than byte equality", () => {
    const input = buildRoundTripFixture();
    const output = writePptx(readPptx(input));

    expect(output).not.toEqual(input);
    expect(readPptx(output).presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
  });

  it("Even in real fixtures, PPTX after write can be reread with readPptx", () => {
    const fixturePath = fileURLToPath(
      new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
    );
    const source = readPptx(readFileSync(fixturePath));
    const output = writePptx(source);
    const reread = readPptx(output);
    const originalImage = source.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );
    const rereadImage = reread.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );

    expect(reread.presentation.slidePartPaths).toEqual(source.presentation.slidePartPaths);
    expect(reread.presentation.slideSize).toEqual(source.presentation.slideSize);
    expect(rereadImage?.bytes).toEqual(originalImage?.bytes);
  });

  it("Parts without preserved material are not implicitly regenerated by no-edit writer.", () => {
    const source = readPptx(buildRoundTripFixture());
    const withoutRawSlide = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        rawParts: source.packageGraph.rawParts?.filter(
          (part) => part.partPath !== "ppt/slides/slide1.xml",
        ),
      },
    };

    expect(() => writePptx(withoutRawSlide)).toThrow(/no preserved package material/);
  });
});

describe("writePptx - slide topology edits", () => {
  it("adds an empty slide that references a layout unused by existing slides", () => {
    const source = readPptx(buildUnreferencedLayoutFixture());
    const targetLayout = source.slideLayouts.find(
      (layout) => layout.partPath === "ppt/slideLayouts/slideLayout3.xml",
    );
    const edited = addEmptySlideFromLayout(source, {
      layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout3.xml"),
    });
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));
    const newSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide2.xml"));
    const newSlideRels = decoder.decode(getEntry(output, "ppt/slides/_rels/slide2.xml.rels"));

    expect(targetLayout).toBeDefined();
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(reread.slides[1]?.layoutPartPath).toBe("ppt/slideLayouts/slideLayout3.xml");
    expect(presentationXml).toContain(`<p:sldId id="257" r:id="rId3"/>`);
    expect(presentationRels).toContain(
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>`,
    );
    expect(newSlideRels).toContain(
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml"/>`,
    );
    expect(newSlideXml).toContain("<p:cSld><p:spTree>");
    expect(newSlideXml).not.toContain("<p:sp>");
    expect(entries["ppt/slides/slide2.xml"]).toBeDefined();
    expect(entries["ppt/slides/_rels/slide2.xml.rels"]).toBeDefined();
    expect(reread.packageGraph.contentTypes.overrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/slide2.xml",
          contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        },
      ]),
    );
  });

  it("duplicates a slide immediately after the source and preserves raw invisible slide material", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = duplicateSlide(source, source.slides[0].handle!);
    const output = writePptx(edited);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));
    const duplicateSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide3.xml"));
    const duplicateSlideRels = decoder.decode(getEntry(output, "ppt/slides/_rels/slide3.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(
      reread.slides.map((slide) => slide.shapes[0]?.kind === "shape" && slide.shapes[0].name),
    ).toEqual(["Invisible Source", "Invisible Source", "Second"]);
    expect(presentationXml).toContain(`<p:sldId id="301" r:id="rId10"/>`);
    expect(presentationXml.indexOf(`r:id="rIdSlide1"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rId10"`),
    );
    expect(presentationXml.indexOf(`r:id="rId10"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rIdSlide2"`),
    );
    expect(presentationRels).toContain(
      `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>`,
    );
    expect(duplicateSlideXml).toContain("<p:timing>");
    expect(duplicateSlideRels).toContain(`Id="rIdComments"`);
    expect(duplicateSlideRels).toContain(`Target="../comments/comment1.xml"`);
    expect(duplicateSlideRels).toContain(`Id="rIdNotes"`);
    expect(duplicateSlideRels).toContain(`Target="../notesSlides/notesSlide2.xml"`);
    expect(decoder.decode(getEntry(output, "ppt/notesSlides/notesSlide2.xml"))).toContain(
      "<p:notes",
    );
    expect(
      decoder.decode(getEntry(output, "ppt/notesSlides/_rels/notesSlide2.xml.rels")),
    ).toContain(`notesMaster`);
    expect(
      decoder.decode(getEntry(output, "ppt/notesSlides/_rels/notesSlide2.xml.rels")),
    ).toContain(`Target="../slides/slide3.xml"`);
    expect(reread.packageGraph.contentTypes.overrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/slide3.xml",
          contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        },
        {
          partName: "ppt/notesSlides/notesSlide2.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        },
      ]),
    );
  });

  it("moves an existing slide by reordering presentation slide ids only", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = moveSlide(source, source.slides[0].handle!, { toIndex: 1 });
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(
      reread.slides.map((slide) => slide.shapes[0]?.kind === "shape" && slide.shapes[0].name),
    ).toEqual(["Second", "Invisible Source"]);
    expect(presentationXml.indexOf(`r:id="rIdSlide2"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rIdSlide1"`),
    );
    expect(presentationRels).toContain(`Id="rIdSlide1"`);
    expect(presentationRels).toContain(`Id="rIdSlide2"`);
    expect(entries["ppt/slides/slide1.xml"]).toBeDefined();
    expect(entries["ppt/slides/slide2.xml"]).toBeDefined();
    expect(entries["ppt/notesSlides/notesSlide1.xml"]).toBeDefined();
  });

  it("deletes a slide and its notes part while keeping remaining slide order and orphan cleanup out of scope", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = deleteSlide(source, source.slides[0].handle!);
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
    expect(reread.slides[0]?.shapes[0]?.kind === "shape" && reread.slides[0].shapes[0].name).toBe(
      "Second",
    );
    expect(presentationXml).not.toContain(`r:id="rIdSlide1"`);
    expect(presentationRels).not.toContain(`Id="rIdSlide1"`);
    expect(entries["ppt/slides/slide1.xml"]).toBeUndefined();
    expect(entries["ppt/slides/_rels/slide1.xml.rels"]).toBeUndefined();
    expect(entries["ppt/notesSlides/notesSlide1.xml"]).toBeUndefined();
    expect(entries["ppt/notesSlides/_rels/notesSlide1.xml.rels"]).toBeUndefined();
    expect(entries["ppt/comments/comment1.xml"]).toBeDefined();
    expect(
      reread.packageGraph.contentTypes.overrides.some(
        (override) => override.partName === "ppt/notesSlides/notesSlide1.xml",
      ),
    ).toBe(false);
  });

  it("keeps relationship content type overrides consistent when no rels default exists", () => {
    const source = readPptx(buildSlideTopologyFixtureWithRelationshipOverrides());
    const duplicated = readPptx(writePptx(duplicateSlide(source, source.slides[0].handle!)));
    const deleted = readPptx(writePptx(deleteSlide(source, source.slides[0].handle!)));
    const duplicatedOverrides = duplicated.packageGraph.contentTypes.overrides;
    const deletedOverridePartNames = new Set(
      deleted.packageGraph.contentTypes.overrides.map((override) => override.partName),
    );

    expect(
      duplicated.packageGraph.contentTypes.defaults.some((entry) => entry.extension === "rels"),
    ).toBe(false);
    expect(duplicatedOverrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/_rels/slide3.xml.rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
        {
          partName: "ppt/notesSlides/_rels/notesSlide2.xml.rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
      ]),
    );
    expect(deletedOverridePartNames.has("ppt/slides/slide1.xml")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/slides/_rels/slide1.xml.rels")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/notesSlides/notesSlide1.xml")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/notesSlides/_rels/notesSlide1.xml.rels")).toBe(false);
  });

  it("rejects deleting the last slide and duplicating a dirty slide", () => {
    const singleSlide = readPptx(buildTextEditFixture());
    expect(() => deleteSlide(singleSlide, singleSlide.slides[0].handle!)).toThrow(/last slide/);

    const dirty = replaceTextRunPlainText(singleSlide, firstRun(singleSlide).handle!, "Dirty");
    expect(() => duplicateSlide(dirty, dirty.slides[0].handle!)).toThrow(
      /pending dirty part edits/,
    );
  });

  it("does not reuse a deleted slide part name within the same edit journal", () => {
    const source = readPptx(buildRoundTripFixture());
    const deletedSecond = deleteSlide(source, source.slides[1].handle!);
    const duplicatedFirst = duplicateSlide(deletedSecond, deletedSecond.slides[0].handle!);
    const deletedDuplicate = deleteSlide(duplicatedFirst, duplicatedFirst.slides[1].handle!);
    const output = writePptx(deletedDuplicate);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));

    expect(duplicatedFirst.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
    ]);
    expect(reread.presentation.slidePartPaths).toEqual(["ppt/slides/slide1.xml"]);
    expect(presentationXml).not.toContain(`r:id="rIdSlide2"`);
    expect(presentationXml).not.toContain(`r:id="rId3"`);
  });

  it("assigns new slide numeric ids at edit time and the writer only applies them", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const duplicatedOnce = duplicateSlide(source, source.slides[0].handle!);
    const duplicatedTwice = duplicateSlide(duplicatedOnce, duplicatedOnce.slides[0].handle!);
    const newSlideNumericIds = (duplicatedTwice.edits ?? []).flatMap((edit) =>
      edit.kind === "duplicateSlide" ? [edit.newSlideNumericId] : [],
    );
    const presentationXml = decoder.decode(
      getEntry(writePptx(duplicatedTwice), "ppt/presentation.xml"),
    );

    expect(newSlideNumericIds).toEqual([301, 302]);
    expect(presentationXml).toContain(`<p:sldId id="301"`);
    expect(presentationXml).toContain(`<p:sldId id="302"`);
  });
});

describe("writePptx - shape add/delete edits", () => {
  it("adds a text box with a collision-free shape id and persists it", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      text: "Added text box",
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const added = requireShape(findShapeByName(reread, "TextBox 31"));
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(added).toMatchObject({
      nodeId: "31",
      name: "TextBox 31",
      transform: {
        offsetX: 914400,
        offsetY: 457200,
        width: 2743200,
        height: 914400,
      },
    });
    expect(added.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Added text box");
    expect(slideXml).toContain(`<p:cNvPr id="31" name="TextBox 31"`);
    expect(slideXml).toContain(`<p:cNvSpPr txBox="1"`);
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toContain("preserve-me");
  });

  it("adds a connector with connection sites, preset geometry, and arrow endpoints", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const start = findShapeByName(source, "Delete Me");
    const end = findShapeByName(source, "Keep Shape");
    const edited = addConnector(source, source.slides[0].handle!, {
      preset: "bentConnector3",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      start: {
        shapeHandle: requireHandle(start.handle),
        connectionSiteIndex: 1,
      },
      end: {
        shapeHandle: requireHandle(end.handle),
        connectionSiteIndex: 3,
      },
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const added = findConnectorByName(reread, "Connector 31");
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(added).toMatchObject({
      nodeId: "31",
      name: "Connector 31",
      connection: {
        start: { shapeId: "10", connectionSiteIndex: 1 },
        end: { shapeId: "30", connectionSiteIndex: 3 },
      },
      geometry: { preset: "bentConnector3" },
      outline: {
        fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
    });
    expect(findConnectorByName(edited, "Connector 31").outline).toMatchObject({
      fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
      tailEnd: { type: "triangle", width: "med", length: "lg" },
    });
    expect(slideXml).toContain(`<p:cxnSp>`);
    expect(slideXml).toContain(`<a:stCxn id="10" idx="1"`);
    expect(slideXml).toContain(`<a:endCxn id="30" idx="3"`);
    expect(slideXml).toContain(`<a:prstGeom prst="bentConnector3"`);
    expect(slideXml).toContain(`<a:tailEnd type="triangle" w="med" len="lg"`);
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toContain("preserve-me");
  });

  it("adds and deletes a free connector without native connection sites", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addConnector(source, source.slides[0].handle!, {
      preset: "straightConnector1",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const added = findConnectorByName(reread, "Connector 31");
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(added).toMatchObject({
      nodeId: "31",
      name: "Connector 31",
      geometry: { preset: "straightConnector1" },
      outline: { tailEnd: { type: "triangle", width: "med", length: "med" } },
    });
    expect(added.connection).toBeUndefined();
    expect(slideXml).toContain(`<p:cxnSp>`);
    expect(slideXml).not.toContain(`<a:stCxn`);
    expect(slideXml).not.toContain(`<a:endCxn`);

    const persisted = readPptx(output);
    const deleted = deleteShape(
      persisted,
      requireHandle(findConnectorByName(persisted, "Connector 31").handle),
    );
    const deletedOutput = writePptx(deleted);
    expect(findConnectorByNameOptional(readPptx(deletedOutput), "Connector 31")).toBeUndefined();
  });

  it("rejects deleting a shape referenced by a connector", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const start = findShapeByName(source, "Delete Me");
    const end = findShapeByName(source, "Keep Shape");
    const withConnector = addConnector(source, source.slides[0].handle!, {
      preset: "straightConnector1",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      start: {
        shapeHandle: requireHandle(start.handle),
        connectionSiteIndex: 1,
      },
      end: {
        shapeHandle: requireHandle(end.handle),
        connectionSiteIndex: 3,
      },
    });

    expect(() => deleteShape(withConnector, requireHandle(start.handle))).toThrow(
      /referenced by connector/,
    );
  });

  it("allows an added text box to be edited before write", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(1000),
      offsetY: asEmu(2000),
      width: asEmu(3000),
      height: asEmu(4000),
      text: "Initial",
      name: "Editable Added",
    });
    const added = requireShape(findShapeByName(withTextBox, "Editable Added"));
    const runHandle = added.textBody?.paragraphs[0]?.runs[0]?.handle;
    if (runHandle === undefined || added.handle === undefined) {
      throw new Error("added text box handles not found");
    }

    const edited = updateShapeTransform(
      replaceTextRunPlainText(withTextBox, runHandle, "Edited Added"),
      added.handle,
      {
        offsetX: asEmu(5000),
        offsetY: asEmu(6000),
        width: asEmu(7000),
        height: asEmu(8000),
      },
    );
    const rereadAdded = requireShape(
      findShapeByName(readPptx(writePptx(edited)), "Editable Added"),
    );

    expect(rereadAdded.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Edited Added");
    expect(rereadAdded.transform).toMatchObject({
      offsetX: 5000,
      offsetY: 6000,
      width: 7000,
      height: 8000,
    });
  });

  it("does not reuse a pending-deleted shape id when adding a text box", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deletedMaxIdShape = deleteShape(
      source,
      requireHandle(findShapeByName(source, "Keep Shape").handle),
    );
    const edited = addTextBox(deletedMaxIdShape, deletedMaxIdShape.slides[0].handle!, {
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      text: "Added after delete",
    });
    const output = writePptx(edited);
    const reread = readPptx(output);

    expect(findShapeByName(reread, "TextBox 31").textBody?.paragraphs[0]?.runs[0]?.text).toBe(
      "Added after delete",
    );
    expect(() => findShapeByName(reread, "Keep Shape")).toThrow(/shape not found/);
  });

  it("does not reuse a pending-deleted shape id when adding a picture", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deletedMaxIdShape = deleteShape(
      source,
      requireHandle(findShapeByName(source, "Keep Shape").handle),
    );
    const edited = addPicture(deletedMaxIdShape, deletedMaxIdShape.slides[0].handle!, {
      bytes: RED_PNG,
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const picture = reread.slides[0]?.shapes.find(
      (shape) => shape.kind === "image" && shape.name === "Picture 31",
    );

    expect(picture).toMatchObject({
      kind: "image",
      nodeId: "31",
      name: "Picture 31",
      blipRelationshipId: "rId1",
    });
    expect(() => findShapeByName(reread, "Keep Shape")).toThrow(/shape not found/);
  });

  it("cancels the add edit when a newly-added text box is deleted before write", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      text: "Temporary",
      name: "Temporary TextBox",
    });
    const added = findShapeByName(withTextBox, "Temporary TextBox");
    const edited = deleteShape(withTextBox, requireHandle(added.handle));
    const output = writePptx(edited);

    expect(
      edited.edits?.filter((edit) => edit.kind === "addTextBox" || edit.kind === "deleteShape"),
    ).toEqual([]);
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).not.toContain(
      "Temporary TextBox",
    );
  });

  it("finalizes added shape XML on the edit record and the writer only splices it", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const start = findShapeByName(source, "Delete Me");
    const end = findShapeByName(source, "Keep Shape");
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      text: "Added text box",
    });
    const edited = addConnector(withTextBox, withTextBox.slides[0].handle!, {
      preset: "straightConnector1",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      start: { shapeHandle: requireHandle(start.handle), connectionSiteIndex: 1 },
      end: { shapeHandle: requireHandle(end.handle), connectionSiteIndex: 3 },
    });
    const textBoxEdit = edited.edits?.find((edit) => edit.kind === "addTextBox");
    const connectorEdit = edited.edits?.find((edit) => edit.kind === "addConnector");
    if (textBoxEdit?.kind !== "addTextBox" || connectorEdit?.kind !== "addConnector") {
      throw new Error("expected addTextBox and addConnector edits to be recorded");
    }
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(textBoxEdit.xml).toContain(`<p:cNvPr id="31" name="TextBox 31"/>`);
    expect(textBoxEdit.xml).toContain(`<a:t>Added text box</a:t>`);
    expect(connectorEdit.xml).toContain(`<a:stCxn id="10" idx="1"/>`);
    expect(slideXml).toContain(textBoxEdit.xml);
    expect(slideXml).toContain(connectorEdit.xml);
  });

  it("round-trips added text box text that needs XML escaping and space preservation", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const text = ` A & B <C> "quoted" `;
    const edited = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      text,
      name: "Escaped TextBox",
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(
      requireShape(findShapeByName(edited, "Escaped TextBox")).textBody?.paragraphs[0]?.runs[0]
        ?.text,
    ).toBe(text);
    expect(
      requireShape(findShapeByName(reread, "Escaped TextBox")).textBody?.paragraphs[0]?.runs[0]
        ?.text,
    ).toBe(text);
    expect(slideXml).toContain(`xml:space="preserve"`);
    expect(slideXml).toContain(`A &amp; B &lt;C&gt;`);
  });

  it("deletes only the targeted sp shape while preserving other shapes and invisible slide material", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deleted = deleteShape(source, requireHandle(source.slides[0].shapes[0]?.handle));
    const output = writePptx(deleted);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);
    const rereadShapeNames: (string | undefined)[] = [];
    for (const shape of reread.slides[0].shapes) {
      rereadShapeNames.push(shape.kind === "raw" ? undefined : shape.name);
    }

    expect(rereadShapeNames).toEqual(expect.arrayContaining(["Keep Picture", "Keep Shape"]));
    expect(reread.slides[0].shapes).toHaveLength(2);
    expect(slideXml).not.toContain("Delete Me");
    expect(slideXml).toContain("Keep Picture");
    expect(slideXml).toContain("Keep Shape");
    expect(slideXml).toContain("<p:timing>");
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toContain("preserve-me");
  });

  it("rejects deleting pic and graphicFrame nodes through the sp/cxnSp delete API", () => {
    const source = readPptx(buildShapeDeleteFixture());

    expect(() => deleteShape(source, requireHandle(source.slides[0].shapes[1]?.handle))).toThrow(
      /only top-level sp or cxnSp shapes/,
    );
  });

  it("rejects conflicting shape additions for the same shape id", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      text: "Added once",
    });
    const additions = withTextBox.edits?.filter((edit) => edit.kind === "addTextBox") ?? [];
    const conflicted = { ...withTextBox, edits: [...(withTextBox.edits ?? []), ...additions] };

    expect(() => writePptx(conflicted)).toThrow(/conflicting shape additions/);
  });

  it("rejects conflicting shape delete edits for the same handle", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deleted = deleteShape(source, requireHandle(source.slides[0].shapes[0]?.handle));
    const deletes = deleted.edits?.filter((edit) => edit.kind === "deleteShape") ?? [];
    const conflicted = { ...deleted, edits: [...(deleted.edits ?? []), ...deletes] };

    expect(() => writePptx(conflicted)).toThrow(/conflicting shape delete edits/);
  });
});

describe("writePptx - one plain text-run edit", () => {
  it("keeps numeric-like text strings in the source and computed view", () => {
    const source = readPptx(buildNumericLikeTextFixture());
    const shape = findShapeByName(source, "Numeric Text");
    const computed = createComputedView(source);
    const computedShape = computed.slides[0]?.elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Numeric Text",
    );

    expect(shape.textBody?.paragraphs[0]?.runs.map((run) => run.text)).toEqual([
      "007",
      "1e5",
      "12.50",
    ]);
    expect(
      computedShape?.kind === "shape"
        ? computedShape.textBody?.paragraphs[0]?.runs.map((run) => run.text)
        : undefined,
    ).toEqual(["007", "1e5", "12.50"]);
  });

  it("Existing text run can be identified with stable source handle", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shape:10:p:0:r:0",
      orderingSlot: 0,
    });
    expect(findTextRunBySourceHandle(source, run.handle!)).toBe(run);
  });

  it("Apply plain text replacement to PptxSourceModel source and reflect in PPTX after write", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const run = firstRun(source);

    const edited = replaceTextRunPlainText(source, run.handle!, "Edited text");
    const reread = readPptx(writePptx(edited));
    const editedRun = firstRun(reread);

    expect(run.text).toBe("Original");
    expect(firstRun(edited).text).toBe("Edited text");
    expect(editedRun.text).toBe("Edited text");
    expect(firstParagraph(reread).runs[1].text).toBe(" Keep ");
  });

  it("does not transform unrelated numeric-like runs when writing a dirty slide", () => {
    const source = readPptx(buildNumericLikeTextFixture());
    const editRun = findShapeByName(source, "Edit Me").textBody?.paragraphs[0]?.runs[0];
    if (editRun?.handle === undefined) throw new Error("edit run handle not found");

    const edited = replaceTextRunPlainText(source, editRun.handle, "Dirty");
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(slideXml).toContain("<a:t>007</a:t>");
    expect(slideXml).toContain("<a:t>1e5</a:t>");
    expect(slideXml).toContain("<a:t>12.50</a:t>");
    expect(
      findShapeByName(reread, "Numeric Text").textBody?.paragraphs[0]?.runs.map((run) => run.text),
    ).toEqual(["007", "1e5", "12.50"]);
    expect(findShapeByName(reread, "Edit Me").textBody?.paragraphs[0]?.runs[0]?.text).toBe("Dirty");
  });

  it("Writes dirty slide XML with a single XML declaration", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = replaceTextRunPlainText(source, firstRun(source).handle!, "Edited text");
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(slideXml.match(/<\?xml/g)).toHaveLength(1);
  });

  it("Preserving run / paragraph formatting and unrelated package material", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const edited = replaceTextRunPlainText(source, firstRun(source).handle!, "Edited text");
    const output = writePptx(edited);
    const reread = readPptx(output);
    const run = firstRun(reread);
    const paragraph = firstParagraph(reread);

    expect(paragraph.properties).toEqual({ align: "center" });
    expect(run.properties).toMatchObject({
      bold: true,
      italic: true,
      fontSize: 24,
      typeface: "Aptos",
      typefaceEa: "Noto Sans JP",
      color: { kind: "srgb", hex: "FF0000" },
    });
    expect(run.rawSidecars?.map((sidecar) => sidecar.node.name) ?? []).not.toContain("a:ea");
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Runs without shape ids can also be reflected in dirty slide XML using shapeSlot handle.", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shapeSlot:1:p:0:r:0",
      orderingSlot: 0,
    });

    const edited = replaceTextRunPlainText(source, run.handle!, "Edited via slot");
    expect(firstRun(readPptx(writePptx(edited))).text).toBe("Edited via slot");
  });
});

describe("writePptx - multiple text edits", () => {
  it("Applies multiple text-run replacements across different shapes and paragraphs in one write.", () => {
    const source = readPptx(buildMultipleTextEditFixture());
    const firstShapeSecondParagraphRun = shapeAt(source, 0).textBody!.paragraphs[1].runs[0];
    const secondShapeRun = shapeAt(source, 1).textBody!.paragraphs[0].runs[0];

    const edited = replaceTextRunPlainText(
      replaceTextRunPlainText(source, firstShapeSecondParagraphRun.handle!, "Edited paragraph"),
      secondShapeRun.handle!,
      "Edited other shape",
    );
    const reread = readPptx(writePptx(edited));

    expect(shapeAt(reread, 0).textBody!.paragraphs[0].runs[0].text).toBe("First paragraph");
    expect(shapeAt(reread, 0).textBody!.paragraphs[1].runs[0].text).toBe("Edited paragraph");
    expect(shapeAt(reread, 1).textBody!.paragraphs[0].runs[0].text).toBe("Edited other shape");
  });

  it("Rejects conflicting duplicate edits for the same text run.", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);
    const edited = replaceTextRunPlainText(
      replaceTextRunPlainText(source, run.handle!, "First edit"),
      run.handle!,
      "Second edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run edits/);
  });

  it("Rejects conflicting text-run and paragraph replacements for the same paragraph.", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);
    const edited = replaceParagraphPlainText(
      replaceTextRunPlainText(source, paragraph.runs[0].handle!, "Run edit"),
      paragraph.handle!,
      "Paragraph edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run and paragraph edits/);
  });
});

describe("writePptx - text run property edits", () => {
  it("Sets all supported text run properties and persists them after write/read", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    const edited = setTextRunProperties(source, run.handle!, {
      bold: false,
      italic: false,
      underline: true,
      fontSize: asPt(32),
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Liberation Sans",
    });
    const reread = readPptx(writePptx(edited));
    const editedRun = firstRun(reread);

    expect(firstRun(edited).properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 32,
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Liberation Sans",
    });
    expect(editedRun.properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 32,
      color: { kind: "srgb", hex: "00AA44" },
      typeface: "Liberation Sans",
      typefaceEa: "Noto Sans JP",
    });
  });

  it("Clears supported text run properties without removing unrelated rPr attributes or children", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="70" name="Decorated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr lang="en-US" b="1" i="1" u="sng" sz="2400" spc="120">` +
          `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill>` +
          `<a:latin typeface="Aptos"/><a:effectLst/><a:ea typeface="Noto Sans JP"/>` +
          `</a:rPr><a:t>Original</a:t></a:r>` +
          `<a:r><a:rPr b="1"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>Untouched</a:t></a:r>` +
          `</a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    const edited = clearTextRunProperties(source, firstRun(source).handle!, [
      "bold",
      "italic",
      "underline",
      "fontSize",
      "color",
      "typeface",
    ]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    const properties = firstRun(reread).properties;
    expect(properties).toMatchObject({ typefaceEa: "Noto Sans JP" });
    expect(properties?.bold).toBeUndefined();
    expect(properties?.italic).toBeUndefined();
    expect(properties?.underline).toBeUndefined();
    expect(properties?.fontSize).toBeUndefined();
    expect(properties?.color).toBeUndefined();
    expect(properties?.typeface).toBeUndefined();
    expect(slideXml).toContain('lang="en-US"');
    expect(slideXml).toContain('spc="120"');
    expect(slideXml).toContain("<a:effectLst");
    expect(slideXml).toContain('<a:ea typeface="Noto Sans JP"');
    expect(firstParagraph(reread).runs[1].properties).toMatchObject({
      bold: true,
      color: { kind: "srgb", hex: "00FF00" },
    });
  });

  it("Creates rPr when setting properties on a run without existing properties", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const edited = setTextRunProperties(source, firstRun(source).handle!, {
      bold: true,
      fontSize: asPt(18),
      color: { kind: "srgb", hex: "112233" },
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml.indexOf("<a:rPr")).toBeLessThan(slideXml.indexOf("<a:t>Original</a:t>"));
    expect(firstRun(readPptx(output)).properties).toMatchObject({
      bold: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "112233" },
    });
  });

  it("Does not dirty a run when clearing properties that are already absent", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const edited = clearTextRunProperties(source, firstRun(source).handle!, ["bold"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(edited).toBe(source);
    expect(slideXml).not.toContain("<a:rPr");
    expect(firstRun(readPptx(output)).properties).toBeUndefined();
  });

  it("Removes an empty latin run property element when clearing typeface", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="71" name="Typeface"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr><a:latin typeface="Aptos"/></a:rPr><a:t>Original</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const edited = clearTextRunProperties(source, firstRun(source).handle!, ["typeface"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).not.toContain("<a:latin");
    expect(slideXml).not.toContain("<a:rPr");
    expect(firstRun(readPptx(output)).properties).toBeUndefined();
  });

  it("Rejects invalid direct text run property helper input", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstRun(source).handle!;

    expect(() => setTextRunProperties(source, handle, { fontSize: asPt(0) })).toThrow(
      /fontSize must be a finite positive pt value/,
    );
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      setTextRunProperties(source, handle, { strikethrough: true }),
    ).toThrow(/unsupported text run property 'strikethrough'/);
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      clearTextRunProperties(source, handle, ["strikethrough"]),
    ).toThrow(/unsupported text run property 'strikethrough'/);
  });

  it("Rejects no-op text run property edits constructed directly in an edit journal", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = {
      ...source,
      edits: [
        {
          kind: "updateTextRunProperties",
          handle: firstRun(source).handle!,
        },
      ],
    } satisfies typeof source;

    expect(() => writePptx(edited)).toThrow(/must set or clear at least one property/);
  });

  it("Applies text and property edits to the same run in one write", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstRun(source).handle!;
    const edited = setTextRunProperties(
      replaceTextRunPlainText(source, handle, "Edited property text"),
      handle,
      { underline: true, color: { kind: "srgb", hex: "336699" } },
    );
    const reread = readPptx(writePptx(edited));

    expect(firstRun(reread).text).toBe("Edited property text");
    expect(firstRun(reread).properties).toMatchObject({
      underline: true,
      color: { kind: "srgb", hex: "336699" },
    });
  });

  it("Rejects conflicting text run property and paragraph replacements for the same paragraph", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);
    const edited = replaceParagraphPlainText(
      setTextRunProperties(source, paragraph.runs[0].handle!, { bold: false }),
      paragraph.handle!,
      "Paragraph edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run properties and paragraph edits/);
  });
});

describe("writePptx - shape fill and outline editing", () => {
  it("Writes solid shape fill and outline color or width edits.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const edited = setShapeOutline(
      setShapeFill(source, shape.handle!, {
        kind: "solid",
        color: { kind: "srgb", hex: "00aa44" },
      }),
      shape.handle!,
      {
        width: asEmu(25400),
        fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
      },
    );
    const output = writePptx(edited);
    const reread = readPptx(output);
    const rereadShape = findShapeByName(reread, "Styled");
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(rereadShape.fill).toEqual({
      kind: "solid",
      color: { kind: "srgb", hex: "00AA44" },
    });
    expect(rereadShape.outline).toMatchObject({
      width: 25400,
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
      dashStyle: "dash",
    });
    expect(slideXml).toContain(`<a:srgbClr val="00AA44"`);
    expect(slideXml).toContain(`<a:ln w="25400">`);
    expect(slideXml).toContain(`<a:prstDash val="dash"`);
  });

  it("Writes noFill for shape fill and connector outline while preserving connector details.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const connector = findConnectorByName(source, "Connector");
    const edited = setShapeOutline(
      setShapeFill(source, shape.handle!, { kind: "none" }),
      connector.handle!,
      { fill: { kind: "none" } },
    );
    const reread = readPptx(writePptx(edited));

    expect(findShapeByName(reread, "Styled").fill).toEqual({ kind: "none" });
    expect(findConnectorByName(reread, "Connector").outline).toMatchObject({
      width: 12700,
      fill: { kind: "none" },
      tailEnd: { type: "triangle", width: "med", length: "med" },
    });
  });

  it("Writes repeated direct helper shape style changes as compact final edits.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const edited = setShapeOutline(
      setShapeFill(
        setShapeOutline(
          setShapeFill(source, shape.handle!, {
            kind: "solid",
            color: { kind: "srgb", hex: "00aa44" },
          }),
          shape.handle!,
          { fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } } },
        ),
        shape.handle!,
        { kind: "none" },
      ),
      shape.handle!,
      { width: asEmu(38100) },
    );
    const rereadShape = findShapeByName(readPptx(writePptx(edited)), "Styled");

    expect(edited.edits).toHaveLength(2);
    expect(rereadShape.fill).toEqual({ kind: "none" });
    expect(rereadShape.outline).toMatchObject({
      width: 38100,
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
    });
  });

  it("Rejects conflicting shape fill and outline edit journals.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const withConflictingFillEdits = {
      ...source,
      edits: [
        { kind: "updateShapeFill", handle: shape.handle!, fill: { kind: "none" } },
        {
          kind: "updateShapeFill",
          handle: shape.handle!,
          fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
        },
      ],
    } satisfies typeof source;
    const withConflictingOutlineEdits = {
      ...source,
      edits: [
        { kind: "updateShapeOutline", handle: shape.handle!, outline: { width: asEmu(1) } },
        { kind: "updateShapeOutline", handle: shape.handle!, outline: { width: asEmu(2) } },
      ],
    } satisfies typeof source;

    expect(() => writePptx(withConflictingFillEdits)).toThrow(/conflicting shape fill edits/);
    expect(() => writePptx(withConflictingOutlineEdits)).toThrow(/conflicting shape outline edits/);
  });
});

describe("writePptx - paragraph property edits", () => {
  it("Sets paragraph alignment, bullet, and level and persists them after write/read", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);

    const edited = setParagraphProperties(source, paragraph.handle!, {
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(firstParagraph(edited).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(firstParagraph(reread).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(slideXml).toContain('<a:pPr algn="r" lvl="2">');
    expect(slideXml).toContain('<a:buChar char="\u2022"');
  });

  it("Writes explicit buNone when removing bullets and supports auto-number bullets", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="72" name="Paragraph props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>Bullet</a:t></a:r></a:p>` +
          `<a:p><a:r><a:t>Numbered</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    const first = firstParagraph(source);
    const second = firstShape(source).textBody!.paragraphs[1];
    const edited = setParagraphProperties(
      setParagraphProperties(source, first.handle!, { bullet: { type: "none" } }),
      second.handle!,
      { level: 1, bullet: { type: "autoNum", scheme: "alphaLcParenR", startAt: 3 } },
    );
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(firstShape(reread).textBody!.paragraphs[0].properties).toMatchObject({
      bullet: { type: "none" },
    });
    expect(firstShape(reread).textBody!.paragraphs[1].properties).toMatchObject({
      level: 1,
      bullet: { type: "autoNum", scheme: "alphaLcParenR", startAt: 3 },
    });
    expect(slideXml).toContain("<a:buNone");
    expect(slideXml).toContain('<a:buAutoNum type="alphaLcParenR" startAt="3"');
  });

  it("Clears only requested paragraph properties and preserves unedited paragraphs", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="73" name="Preserve paragraph props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr algn="ctr" lvl="2"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>Edit</a:t></a:r></a:p>` +
          `<a:p><a:pPr algn="r" lvl="1"><a:buChar char="&#x25E6;"/></a:pPr><a:r><a:t>Keep</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const edited = clearParagraphProperties(source, firstParagraph(source).handle!, ["align"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(firstParagraph(reread).properties).toMatchObject({
      level: 2,
      bullet: { type: "char", char: "\u2022" },
      lineSpacing: { type: "pct", value: 90000 },
    });
    expect(firstParagraph(reread).properties?.align).toBeUndefined();
    expect(firstShape(reread).textBody!.paragraphs[1].properties).toMatchObject({
      align: "right",
      level: 1,
      bullet: { type: "char", char: "\u25E6" },
    });
    expect(slideXml).toContain('<a:lnSpc><a:spcPct val="90000"');
    expect(slideXml).toContain('<a:pPr algn="r" lvl="1"');
  });

  it("Rejects no-op paragraph property edits constructed directly in an edit journal", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = {
      ...source,
      edits: [
        {
          kind: "updateParagraphProperties",
          handle: firstParagraph(source).handle!,
        },
      ],
    } satisfies typeof source;

    expect(() => writePptx(edited)).toThrow(/must set or clear at least one property/);
  });
});

describe("writePptx - paragraph text replacement", () => {
  it("Normalizes a multi-run paragraph to one run using the first run properties.", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const paragraph = firstParagraph(source);

    expect(findParagraphBySourceHandle(source, paragraph.handle!)).toBe(paragraph);

    const edited = replaceParagraphPlainText(source, paragraph.handle!, "Paragraph replacement");
    const output = writePptx(edited);
    const reread = readPptx(output);
    const editedParagraph = firstParagraph(reread);

    expect(firstParagraph(edited).runs).toHaveLength(1);
    expect(editedParagraph.runs).toHaveLength(1);
    expect(editedParagraph.runs[0].text).toBe("Paragraph replacement");
    expect(editedParagraph.runs[0].properties).toMatchObject({
      bold: true,
      italic: true,
      fontSize: 24,
      typeface: "Aptos",
      typefaceEa: "Noto Sans JP",
      color: { kind: "srgb", hex: "FF0000" },
    });
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).not.toContain(" Keep ");
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Round-trips paragraph replacement text with significant surrounding whitespace.", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = replaceParagraphPlainText(source, firstParagraph(source).handle!, " Trimmed ");
    const reread = readPptx(writePptx(edited));

    expect(firstParagraph(reread).runs).toHaveLength(1);
    expect(firstRun(reread).text).toBe(" Trimmed ");
  });

  it("Keeps replacement run before endParaRPr.", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="60" name="End para props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr algn="ctr"/><a:r><a:t>Before</a:t></a:r><a:endParaRPr lang="ja-JP"/></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const output = writePptx(
      replaceParagraphPlainText(source, firstParagraph(source).handle!, "After"),
    );
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml.indexOf("<a:r>")).toBeGreaterThan(-1);
    expect(slideXml.indexOf("<a:endParaRPr")).toBeGreaterThan(-1);
    expect(slideXml.indexOf("<a:r>")).toBeLessThan(slideXml.indexOf("<a:endParaRPr"));
    expect(firstRun(readPptx(output)).text).toBe("After");
  });

  it("Rejects paragraph replacement for interleaved bullet paragraphs split by the reader.", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="61" name="Interleaved bullets"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p>` +
          `<a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>One</a:t></a:r><a:br/>` +
          `<a:pPr><a:buChar char="&#x25E6;"/></a:pPr><a:r><a:t>Two</a:t></a:r>` +
          `</a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    expect(firstShape(source).textBody!.paragraphs).toHaveLength(2);
    const edited = replaceParagraphPlainText(source, firstParagraph(source).handle!, "After");

    expect(() => writePptx(edited)).toThrow(/interleaved bullet paragraph/);
  });
});

describe("writePptx - shape xfrm edit", () => {
  it("Apply offset and extent update to PptxSourceModel source and reflect in PPTX after write", () => {
    const source = readPptx(buildTextEditFixture());
    const shape = firstShape(source);
    const handle = shape.handle!;

    const edited = updateShapeTransform(source, handle, {
      offsetX: asEmu(1111),
      offsetY: asEmu(2222),
      width: asEmu(3333),
      height: asEmu(4444),
    });
    const reread = readPptx(writePptx(edited));
    const editedShape = findShapeNodeBySourceHandle(reread, handle);

    expect(firstShape(edited).transform).toMatchObject({
      offsetX: 1111,
      offsetY: 2222,
      width: 3333,
      height: 4444,
    });
    expect(editedShape?.transform).toMatchObject({
      offsetX: 1111,
      offsetY: 2222,
      width: 3333,
      height: 4444,
    });
  });

  it("Rejects conflicting shape transform edits for the same shape", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstShape(source).handle!;
    const edited = updateShapeTransform(
      updateShapeTransform(source, handle, {
        offsetX: asEmu(1111),
        offsetY: asEmu(2222),
        width: asEmu(3333),
        height: asEmu(4444),
      }),
      handle,
      {
        offsetX: asEmu(5555),
        offsetY: asEmu(6666),
        width: asEmu(7777),
        height: asEmu(8888),
      },
    );

    expect(() => writePptx(edited)).toThrow(/conflicting shape transform edits/);
  });

  it("Preserves unrelated package material while replacing only dirty slide XML", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const edited = updateShapeTransform(source, firstShape(source).handle!, {
      offsetX: asEmu(1111),
      offsetY: asEmu(2222),
      width: asEmu(3333),
      height: asEmu(4444),
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).toContain('<a:off x="1111" y="2222"');
    expect(slideXml).toContain('<a:ext cx="3333" cy="4444"');
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Rejects shape handles that do not have xfrm", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="50" name="No xfrm"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`,
      ),
    );
    const shapeWithoutXfrm = firstShape(source);

    expect(() =>
      updateShapeTransform(source, shapeWithoutXfrm.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/does not reference a shape with xfrm/);
  });

  it("Rejects shape transform handles without node ids", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr name="No Id Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`,
      ),
    );
    const shape = firstShape(source);

    expect(shape.handle?.nodeId).toBeUndefined();
    expect(() =>
      updateShapeTransform(source, shape.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/requires a node id/);
  });

  it("Rejects nonexistent shape handles", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = {
      ...firstShape(source).handle!,
      nodeId: asSourceNodeId("999"),
    };

    expect(() =>
      updateShapeTransform(source, handle, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/shape handle was not found/);
  });

  it("Rejects nested group child shape handles for this writer slice", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="30" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/><a:chOff x="0" y="0"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="31" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `</p:sp>` +
          `</p:grpSp>`,
      ),
    );
    const group = source.slides[0].shapes[0];
    if (group.kind !== "group") throw new Error("group not found");
    const child = group.children[0];

    expect(findShapeNodeBySourceHandle(source, child.handle!)).toBe(child);
    expect(() =>
      updateShapeTransform(source, child.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/nested group shape editing is not supported/);
  });

  it("Rejects AlternateContent fallback shape handles for this writer slice", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Fallback>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="40" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `</p:sp>` +
          `</mc:Fallback>` +
          `</mc:AlternateContent>`,
      ),
    );
    const shape = firstShape(source);

    expect(() =>
      updateShapeTransform(source, shape.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/AlternateContent/);
  });
});

function firstShape(source: ReturnType<typeof readPptx>): SourceShape {
  return shapeAt(source, 0);
}

function shapeAt(source: ReturnType<typeof readPptx>, index: number): SourceShape {
  const shape = source.slides[0].shapes.filter(
    (node): node is SourceShape => node.kind === "shape",
  )[index];
  if (shape === undefined) throw new Error("shape not found");
  return shape;
}

function findShapeByName(source: ReturnType<typeof readPptx>, name: string): SourceShape {
  const shape = source.slides
    .flatMap((slide) => slide.shapes)
    .find((node): node is SourceShape => node.kind === "shape" && node.name === name);
  if (shape === undefined) throw new Error(`shape not found: ${name}`);
  return shape;
}

function findImageByName(source: ReturnType<typeof readPptx>, name: string): SourceImage {
  const image = source.slides
    .flatMap((slide) => slide.shapes)
    .find((node): node is SourceImage => node.kind === "image" && node.name === name);
  if (image === undefined) throw new Error(`image not found: ${name}`);
  return image;
}

function findConnectorByName(source: ReturnType<typeof readPptx>, name: string): SourceConnector {
  const connector = findConnectorByNameOptional(source, name);
  if (connector === undefined) throw new Error(`connector not found: ${name}`);
  return connector;
}

function findConnectorByNameOptional(
  source: ReturnType<typeof readPptx>,
  name: string,
): SourceConnector | undefined {
  return source.slides
    .flatMap((slide) => slide.shapes)
    .find((node): node is SourceConnector => node.kind === "connector" && node.name === name);
}

function requireShape(shape: SourceShape | undefined): SourceShape {
  if (shape === undefined) throw new Error("shape not found");
  return shape;
}

function requireHandle(handle: SourceShapeNode["handle"]): NonNullable<SourceShapeNode["handle"]> {
  if (handle === undefined) throw new Error("handle not found");
  return handle;
}

function firstParagraph(source: ReturnType<typeof readPptx>) {
  return firstShape(source).textBody!.paragraphs[0];
}

function firstRun(source: ReturnType<typeof readPptx>) {
  return firstParagraph(source).runs[0];
}
