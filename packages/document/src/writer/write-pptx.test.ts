import { Buffer } from "node:buffer";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  asEmu,
  asPartPath,
  asPt,
  asSourceNodeId,
  createComputedView,
  createPptx,
  readPptx,
  writePptx,
} from "../index.js";
import {
  addConnector,
  addEmptySlideFromLayout,
  addTextBox,
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
  setTextRunProperties,
  type SourceConnector,
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
