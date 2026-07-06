import sharp from "sharp";

import type { FixtureCreatorMap } from "../fixture-builder.js";
import {
  buildPptx,
  gradientFillXml,
  NS,
  REL_TYPES,
  savePptx,
  shapeXml,
  slideRelsXml,
  solidFillXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../fixture-builder.js";

async function createBackgroundFixture(): Promise<void> {
  // Slide 1: Solid background
  const bgSolid = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:bgPr></p:bg>`;
  const slide1Shapes = shapeXml(2, "text-on-bg", {
    preset: "rect",
    x: 2000000,
    y: 1500000,
    cx: 5000000,
    cy: 2000000,
    fillXml: solidFillXml("FFFFFF"),
    textBodyXml: textBodyXmlHelper("Solid Background", {
      fontSize: 24,
      color: "333333",
    }),
  });
  const slide1 = wrapSlideXml(slide1Shapes, bgSolid);

  // Slide 2: Gradient background
  const bgGrad = `<p:bg><p:bgPr>${gradientFillXml(
    [
      { pos: 0, color: "1A1A2E" },
      { pos: 50000, color: "16213E" },
      { pos: 100000, color: "0F3460" },
    ],
    5400000,
  )}</p:bgPr></p:bg>`;
  const slide2Shapes = shapeXml(2, "text-on-grad-bg", {
    preset: "roundRect",
    x: 2000000,
    y: 1500000,
    cx: 5000000,
    cy: 2000000,
    fillXml: `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="80000"/></a:srgbClr></a:solidFill>`,
    textBodyXml: textBodyXmlHelper("Gradient Background", {
      fontSize: 24,
      color: "333333",
    }),
  });
  const slide2 = wrapSlideXml(slide2Shapes, bgGrad);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "background.pptx");
}

async function createBackgroundBlipFillFixture(): Promise<void> {
  // Generate a gradient-like test image for background
  const imgSize = 200;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = Math.floor((x / imgSize) * 255); // R
      pixels[idx + 1] = Math.floor((y / imgSize) * 200); // G
      pixels[idx + 2] = 100; // B
      pixels[idx + 3] = 255; // A
    }
  }
  const bgImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  // Slide master with blipFill background
  const masterWithBgImage = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:blipFill>
          <a:blip r:embed="rId3"/>
          <a:stretch><a:fillRect/></a:stretch>
        </a:blipFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

  const masterRelsWithImage = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="${REL_TYPES.theme}" Target="../theme/theme1.xml"/>
  <Relationship Id="rId3" Type="${REL_TYPES.image}" Target="../media/image1.png"/>
</Relationships>`;

  // Slide with text on top of the background image
  const slideShapes = shapeXml(2, "text-on-bg-image", {
    preset: "roundRect",
    x: 2000000,
    y: 1500000,
    cx: 5000000,
    cy: 2000000,
    fillXml: `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="70000"/></a:srgbClr></a:solidFill>`,
    textBodyXml: textBodyXmlHelper("Background Image from Master", {
      fontSize: 24,
      color: "333333",
    }),
  });
  const slide = wrapSlideXml(slideShapes);
  const rels = slideRelsXml();

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", bgImage);

  const buffer = await buildPptx({
    slides: [{ xml: slide, rels }],
    media,
    slideMasterXml: masterWithBgImage,
    slideMasterRelsXml: masterRelsWithImage,
  });
  savePptx(buffer, "background-blipfill.pptx");
}

export const backgroundFixtureCreators: FixtureCreatorMap = {
  "background.pptx": createBackgroundFixture,
  "background-blipfill.pptx": createBackgroundBlipFillFixture,
};
