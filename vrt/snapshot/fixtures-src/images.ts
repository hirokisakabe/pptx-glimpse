import sharp from "sharp";

import type { FixtureCreatorMap } from "../fixture-builder.js";
import {
  buildPptx,
  gridPosition,
  REL_TYPES,
  savePptx,
  shapeXml,
  slideRelsXml,
  wrapSlideXml,
} from "../fixture-builder.js";

async function createImageFixture(): Promise<void> {
  // Generate a small test image (colored grid)
  const imgSize = 100;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = x < imgSize / 2 ? 255 : 0; // R
      pixels[idx + 1] = y < imgSize / 2 ? 255 : 0; // G
      pixels[idx + 2] = 128; // B
      pixels[idx + 3] = 255; // A
    }
  }
  const testImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  const picXml = `<p:pic>
  <p:nvPicPr><p:cNvPr id="2" name="Image 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="2000000" y="1000000"/><a:ext cx="5000000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;

  const slide = wrapSlideXml(picXml);
  const rels = slideRelsXml([{ id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" }]);

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", testImage);

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }], media });
  savePptx(buffer, "image.pptx");
}

function radialGradientFillXml(
  stops: { pos: number; color: string }[],
  center: { l: number; t: number; r: number; b: number },
): string {
  const gsItems = stops
    .map((s) => `<a:gs pos="${s.pos}"><a:srgbClr val="${s.color}"/></a:gs>`)
    .join("");
  return `<a:gradFill><a:gsLst>${gsItems}</a:gsLst><a:path path="circle"><a:fillToRect l="${center.l}" t="${center.t}" r="${center.r}" b="${center.b}"/></a:path></a:gradFill>`;
}

function patternFillXml(preset: string, fgColor: string, bgColor: string): string {
  return `<a:pattFill prst="${preset}"><a:fgClr><a:srgbClr val="${fgColor}"/></a:fgClr><a:bgClr><a:srgbClr val="${bgColor}"/></a:bgClr></a:pattFill>`;
}

async function createPatternImageFillFixture(): Promise<void> {
  // --- Slide 1: Radial gradients ---
  let id = 2;
  const radialShapes: string[] = [];

  const radialConfigs = [
    {
      name: "RadialCenter",
      preset: "rect",
      col: 0,
      row: 0,
      center: { l: 50000, t: 50000, r: 50000, b: 50000 },
      colors: [
        { pos: 0, color: "FF0000" },
        { pos: 100000, color: "0000FF" },
      ],
    },
    {
      name: "RadialTopLeft",
      preset: "roundRect",
      col: 1,
      row: 0,
      center: { l: 0, t: 0, r: 100000, b: 100000 },
      colors: [
        { pos: 0, color: "FFFF00" },
        { pos: 100000, color: "008000" },
      ],
    },
    {
      name: "RadialBottomRight",
      preset: "ellipse",
      col: 2,
      row: 0,
      center: { l: 100000, t: 100000, r: 0, b: 0 },
      colors: [
        { pos: 0, color: "FFFFFF" },
        { pos: 50000, color: "FFC000" },
        { pos: 100000, color: "FF6384" },
      ],
    },
    {
      name: "RadialRect",
      preset: "rect",
      col: 0,
      row: 1,
      center: { l: 50000, t: 50000, r: 50000, b: 50000 },
      colors: [
        { pos: 0, color: "4472C4" },
        { pos: 100000, color: "ED7D31" },
      ],
    },
  ];

  for (const cfg of radialConfigs) {
    const pos = gridPosition(cfg.col, cfg.row, 3, 2);
    radialShapes.push(
      shapeXml(id++, cfg.name, {
        preset: cfg.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: radialGradientFillXml(cfg.colors, cfg.center),
      }),
    );
  }

  const slide1 = wrapSlideXml(radialShapes.join("\n"));
  const rels1 = slideRelsXml();

  // --- Slide 2: Image fills ---
  id = 2;
  const imgSize = 80;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = Math.floor((x / imgSize) * 255);
      pixels[idx + 1] = Math.floor((y / imgSize) * 255);
      pixels[idx + 2] = 128;
      pixels[idx + 3] = 255;
    }
  }
  const testImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  const imgFillConfigs = [
    { name: "ImageFillRect", preset: "rect", col: 0, row: 0 },
    { name: "ImageFillRoundRect", preset: "roundRect", col: 1, row: 0 },
    { name: "ImageFillEllipse", preset: "ellipse", col: 0, row: 1 },
  ];
  const imgFillShapes: string[] = [];
  for (const cfg of imgFillConfigs) {
    const pos = gridPosition(cfg.col, cfg.row, 2, 2);
    imgFillShapes.push(
      shapeXml(id++, cfg.name, {
        preset: cfg.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: `<a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
      }),
    );
  }

  const slide2 = wrapSlideXml(imgFillShapes.join("\n"));
  const rels2 = slideRelsXml([
    { id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" },
  ]);

  // --- Slide 3: Pattern fills ---
  id = 2;
  const patternPresets = [
    "ltHorz",
    "ltVert",
    "ltDnDiag",
    "ltUpDiag",
    "dkHorz",
    "dkVert",
    "cross",
    "diagCross",
    "pct25",
  ];
  const pattShapes: string[] = [];
  for (let i = 0; i < patternPresets.length; i++) {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 3);
    pattShapes.push(
      shapeXml(id++, `Pattern-${patternPresets[i]}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: patternFillXml(patternPresets[i], "4472C4", "FFFFFF"),
      }),
    );
  }

  const slide3 = wrapSlideXml(pattShapes.join("\n"));
  const rels3 = slideRelsXml();

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", testImage);

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels: rels1 },
      { xml: slide2, rels: rels2 },
      { xml: slide3, rels: rels3 },
    ],
    media,
  });
  savePptx(buffer, "pattern-image-fill.pptx");
}

async function createImageCropFixture(): Promise<void> {
  const imgSize = 100;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = x < imgSize / 2 ? 255 : 0; // R
      pixels[idx + 1] = y < imgSize / 2 ? 255 : 0; // G
      pixels[idx + 2] = 128; // B
      pixels[idx + 3] = 255; // A
    }
  }
  const testImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  // 1. Crop left 25%
  const pic1 = `<p:pic>
  <p:nvPicPr><p:cNvPr id="2" name="CropLeft"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:srcRect l="25000"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="300000" y="300000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;

  // 2. Crop top 20%, bottom 20%
  const pic2 = `<p:pic>
  <p:nvPicPr><p:cNvPr id="3" name="CropTopBottom"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:srcRect t="20000" b="20000"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="4000000" y="300000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;

  // 3. Crop all sides 10%
  const pic3 = `<p:pic>
  <p:nvPicPr><p:cNvPr id="4" name="CropAllSides"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:srcRect l="10000" t="10000" r="10000" b="10000"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="300000" y="2700000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;

  const slide = wrapSlideXml([pic1, pic2, pic3].join("\n"));
  const rels = slideRelsXml([{ id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" }]);

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", testImage);

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }], media });
  savePptx(buffer, "image-crop.pptx");
}

async function createBlipEffectsFixture(): Promise<void> {
  // Generate a small test image (colored grid)
  const imgSize = 100;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = x < imgSize / 2 ? 255 : 0;
      pixels[idx + 1] = y < imgSize / 2 ? 255 : 0;
      pixels[idx + 2] = 128;
      pixels[idx + 3] = 255;
    }
  }
  const testImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  const cols = 4;
  const rows = 2;

  const blipEffects: { name: string; blipXml: string }[] = [
    { name: "Original", blipXml: `<a:blip r:embed="rId2"/>` },
    { name: "Grayscale", blipXml: `<a:blip r:embed="rId2"><a:grayscl/></a:blip>` },
    { name: "BiLevel", blipXml: `<a:blip r:embed="rId2"><a:biLevel thresh="50000"/></a:blip>` },
    { name: "Blur", blipXml: `<a:blip r:embed="rId2"><a:blur rad="50800" grow="0"/></a:blip>` },
    {
      name: "Bright",
      blipXml: `<a:blip r:embed="rId2"><a:lum bright="40000" contrast="0"/></a:blip>`,
    },
    {
      name: "Duotone",
      blipXml: `<a:blip r:embed="rId2"><a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"/></a:duotone></a:blip>`,
    },
    { name: "EMF Placeholder", blipXml: "" },
    { name: "WMF Placeholder", blipXml: "" },
  ];

  let id = 2;
  const shapes: string[] = [];
  const relsExtra: { id: string; type: string; target: string }[] = [
    { id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" },
    { id: "rId3", type: REL_TYPES.image, target: "../media/image2.emf" },
    { id: "rId4", type: REL_TYPES.image, target: "../media/image3.wmf" },
  ];

  for (let i = 0; i < blipEffects.length; i++) {
    const { name, blipXml } = blipEffects[i];
    const col = i % cols;
    const row = Math.floor(i / cols);
    const pos = gridPosition(col, row, cols, rows);

    if (name === "EMF Placeholder") {
      shapes.push(`<p:pic>
  <p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId3"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${pos.x}" y="${pos.y}"/><a:ext cx="${pos.w}" cy="${pos.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`);
    } else if (name === "WMF Placeholder") {
      shapes.push(`<p:pic>
  <p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId4"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${pos.x}" y="${pos.y}"/><a:ext cx="${pos.w}" cy="${pos.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`);
    } else {
      shapes.push(`<p:pic>
  <p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    ${blipXml}
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${pos.x}" y="${pos.y}"/><a:ext cx="${pos.w}" cy="${pos.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`);
    }
    id++;
  }

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml(relsExtra);

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", testImage);
  media.set("ppt/media/image2.emf", Buffer.from("dummy-emf-data"));
  media.set("ppt/media/image3.wmf", Buffer.from("dummy-wmf-data"));

  const buffer = await buildPptx({
    slides: [{ xml: slide, rels }],
    media,
    contentTypesExtra: [
      `<Default Extension="emf" ContentType="image/x-emf"/>`,
      `<Default Extension="wmf" ContentType="image/x-wmf"/>`,
    ],
  });
  savePptx(buffer, "blip-effects.pptx");
}

async function createImageStretchTileFixture(): Promise<void> {
  const imgSize = 100;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = x < imgSize / 2 ? 255 : 0;
      pixels[idx + 1] = y < imgSize / 2 ? 255 : 0;
      pixels[idx + 2] = 128;
      pixels[idx + 3] = 255;
    }
  }
  const testImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  const cols = 3;
  const rows = 1;

  const cases: { name: string; fillXml: string }[] = [
    {
      name: "Stretch Default",
      fillXml: `<a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch>`,
    },
    {
      name: "Stretch Inset",
      fillXml: `<a:blip r:embed="rId2"/><a:stretch><a:fillRect l="15000" t="15000" r="15000" b="15000"/></a:stretch>`,
    },
    {
      name: "Tile 50%",
      fillXml: `<a:blip r:embed="rId2"/><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="none" algn="tl"/>`,
    },
  ];

  let id = 2;
  const shapes: string[] = [];

  for (let i = 0; i < cases.length; i++) {
    const { name, fillXml } = cases[i];
    const pos = gridPosition(i % cols, Math.floor(i / cols), cols, rows);
    shapes.push(`<p:pic>
  <p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>${fillXml}</p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${pos.x}" y="${pos.y}"/><a:ext cx="${pos.w}" cy="${pos.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`);
    id++;
  }

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml([{ id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" }]);

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", testImage);

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }], media });
  savePptx(buffer, "image-stretch-tile.pptx");
}

export const imageFixtureCreators: FixtureCreatorMap = {
  "image.pptx": createImageFixture,
  "pattern-image-fill.pptx": createPatternImageFillFixture,
  "image-crop.pptx": createImageCropFixture,
  "blip-effects.pptx": createBlipEffectsFixture,
  "image-stretch-tile.pptx": createImageStretchTileFixture,
};
