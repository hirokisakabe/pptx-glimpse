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

/**
 * Deterministic metafiles authored from the public Microsoft format specifications:
 * https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/91c257d7-c39d-4a36-9b1f-63e3f73d30ca
 * https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/4813e7fd-52d0-4f42-965f-228c8b7488d2
 */
export function createRepresentativeEmf(): Buffer {
  const records = [
    emfRecord(0x09, int32Values(1000, 1000)),
    emfRecord(0x0b, int32Values(1000, 1000)),
    emfRecord(0x26, uint32Values(1, 0, 8, 0, 0x003366cc)),
    emfRecord(0x27, uint32Values(2, 0, 0x00f2c94c, 0)),
    emfRecord(0x25, uint32Values(1)),
    emfRecord(0x25, uint32Values(2)),
    emfRecord(0x2b, int32Values(80, 80, 920, 600)),
    emfRecord(0x18, uint32Values(0x00333333)),
    emfFontRecord(3, -72, 120, "Noto Sans"),
    emfRecord(0x25, uint32Values(3)),
    emfRecord(0x0a, int32Values(100, 50)),
    emfRecord(0x09, int32Values(1000, 1000)),
    emfRecord(0x0c, int32Values(0, 0)),
    emfRecord(0x0b, int32Values(1000, 1000)),
    emfTextRecord("EMF", 350, 870, [150, 150, 150]),
    emfRecord(0x0e, uint32Values(0, 0, 20)),
  ];
  const totalBytes = 88 + records.reduce((sum, record) => sum + record.length, 0);
  const header = Buffer.alloc(88);
  header.writeUInt32LE(1, 0);
  header.writeUInt32LE(88, 4);
  writeInt32Sequence(header, 8, [0, 0, 1000, 1000, 0, 0, 2540, 2540]);
  header.writeUInt32LE(0x464d4520, 40);
  header.writeUInt32LE(0x00010000, 44);
  header.writeUInt32LE(totalBytes, 48);
  header.writeUInt32LE(records.length + 1, 52);
  header.writeUInt16LE(3, 56);
  header.writeUInt32LE(1000, 72);
  header.writeUInt32LE(1000, 76);
  header.writeUInt32LE(254, 80);
  header.writeUInt32LE(254, 84);
  return Buffer.concat([header, ...records]);
}

export function createRepresentativeWmf(): Buffer {
  const text = Buffer.from("WMF", "latin1");
  const paddedText = text.length % 2 === 0 ? text : Buffer.concat([text, Buffer.alloc(1)]);
  const records = [
    wmfRecord(0x020b, int16Values(0, 0)),
    wmfRecord(0x020c, int16Values(1000, 1000)),
    wmfRecord(
      0x02fa,
      Buffer.concat([uint16Values(0), int16Values(8, 0), uint32Values(0x00cc6633)]),
    ),
    wmfRecord(0x02fc, Buffer.concat([uint16Values(0), uint32Values(0x004cc9f2), uint16Values(0)])),
    wmfRecord(0x012d, uint16Values(0)),
    wmfRecord(0x012d, uint16Values(1)),
    wmfRecord(0x041b, int16Values(620, 920, 80, 80)),
    wmfRecord(0x0209, uint32Values(0x00333333)),
    wmfRecord(
      0x0521,
      Buffer.concat([uint16Values(text.length), paddedText, int16Values(830, 250)]),
    ),
    wmfRecord(0, Buffer.alloc(0)),
  ];
  const recordsBytes = records.reduce((sum, record) => sum + record.length, 0);
  const header = Buffer.alloc(18);
  header.writeUInt16LE(1, 0);
  header.writeUInt16LE(9, 2);
  header.writeUInt16LE(0x0300, 4);
  header.writeUInt32LE((header.length + recordsBytes) / 2, 6);
  header.writeUInt16LE(2, 10);
  header.writeUInt32LE(Math.max(...records.map((record) => record.length / 2)), 12);
  return Buffer.concat([header, ...records]);
}

function emfRecord(type: number, payload: Buffer): Buffer {
  const record = Buffer.alloc(8 + payload.length);
  record.writeUInt32LE(type, 0);
  record.writeUInt32LE(record.length, 4);
  payload.copy(record, 8);
  return record;
}

function emfTextRecord(text: string, x: number, y: number, advances?: readonly number[]): Buffer {
  const encoded = Buffer.from(text, "utf16le");
  const stringEnd = 68 + encoded.length;
  const dxStart = Math.ceil(stringEnd / 4) * 4;
  const payload = Buffer.alloc(dxStart + (advances === undefined ? 0 : advances.length * 4));
  writeInt32Sequence(payload, 0, [0, 0, 1000, 1000]);
  payload.writeUInt32LE(1, 16);
  payload.writeFloatLE(1, 20);
  payload.writeFloatLE(1, 24);
  payload.writeInt32LE(x, 28);
  payload.writeInt32LE(y, 32);
  payload.writeUInt32LE(text.length, 36);
  payload.writeUInt32LE(76, 40);
  writeInt32Sequence(payload, 48, [0, 0, 1000, 1000]);
  if (advances !== undefined) payload.writeUInt32LE(dxStart + 8, 64);
  encoded.copy(payload, 68);
  advances?.forEach((advance, index) => payload.writeInt32LE(advance, dxStart + index * 4));
  return emfRecord(0x54, payload);
}

function emfFontRecord(
  index: number,
  height: number,
  escapement: number,
  faceName: string,
): Buffer {
  const payload = Buffer.alloc(96);
  payload.writeUInt32LE(index, 0);
  payload.writeInt32LE(height, 4);
  payload.writeInt32LE(0, 8);
  payload.writeInt32LE(escapement, 12);
  payload.writeInt32LE(escapement, 16);
  payload.writeInt32LE(700, 20);
  Buffer.from(faceName, "utf16le").copy(payload, 32, 0, 64);
  return emfRecord(0x52, payload);
}

function wmfRecord(type: number, payload: Buffer): Buffer {
  const record = Buffer.alloc(6 + payload.length);
  record.writeUInt32LE(record.length / 2, 0);
  record.writeUInt16LE(type, 4);
  payload.copy(record, 6);
  return record;
}

function uint32Values(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 4);
  values.forEach((value, index) => buffer.writeUInt32LE(value, index * 4));
  return buffer;
}

function int32Values(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 4);
  writeInt32Sequence(buffer, 0, values);
  return buffer;
}

function int16Values(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 2);
  values.forEach((value, index) => buffer.writeInt16LE(value, index * 2));
  return buffer;
}

function uint16Values(...values: number[]): Buffer {
  const buffer = Buffer.alloc(values.length * 2);
  values.forEach((value, index) => buffer.writeUInt16LE(value, index * 2));
  return buffer;
}

function writeInt32Sequence(buffer: Buffer, offset: number, values: readonly number[]): void {
  values.forEach((value, index) => buffer.writeInt32LE(value, offset + index * 4));
}

export async function buildMetafileImagesFixture(): Promise<Buffer> {
  const emf = createRepresentativeEmf();
  const wmf = createRepresentativeWmf();
  const picture = (name: string, relId: string, x: number, crop = "") => `<p:pic>
  <p:nvPicPr><p:cNvPr id="${relId === "rId2" ? 2 : 3}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="${relId}"/>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="${x}" y="350000"/><a:ext cx="3600000" cy="1800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>`;
  const imageFill = `<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="EMF image fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="500000" y="2600000"/><a:ext cx="3600000" cy="1800000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:ln w="38100"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></p:spPr>
</p:sp>`;
  const olePreview = `<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="5" name="OLE WMF preview"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="4700000" y="2600000"/><a:ext cx="3600000" cy="1800000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj r:id="rId4" progId="Package"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="6" name="OLE WMF preview image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="250000" y="250000"/><a:ext cx="1200000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic>
</p:graphicFrame>`;
  const slide = wrapSlideXml(
    [
      picture("EMF cropped picture", "rId2", 500000, `<a:srcRect l="10000"/>`),
      picture("WMF picture", "rId3", 4700000),
      imageFill,
      olePreview,
    ].join("\n"),
  );
  const rels = slideRelsXml([
    { id: "rId2", type: REL_TYPES.image, target: "../media/representative.emf" },
    { id: "rId3", type: REL_TYPES.image, target: "../media/representative.wmf" },
    {
      id: "rId4",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
      target: "../embeddings/oleObject1.bin",
    },
  ]);
  const media = new Map<string, Buffer>([
    ["ppt/media/representative.emf", emf],
    ["ppt/media/representative.wmf", wmf],
    ["ppt/embeddings/oleObject1.bin", Buffer.from("fixture-only OLE payload")],
  ]);
  return buildPptx({
    slides: [{ xml: slide, rels }],
    media,
    contentTypesExtra: [
      `<Default Extension="emf" ContentType="image/x-emf"/>`,
      `<Default Extension="wmf" ContentType="image/x-wmf"/>`,
      `<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>`,
    ],
  });
}

async function createMetafileImagesFixture(): Promise<void> {
  savePptx(await buildMetafileImagesFixture(), "metafile-images.pptx");
}

export const imageFixtureCreators: FixtureCreatorMap = {
  "image.pptx": createImageFixture,
  "pattern-image-fill.pptx": createPatternImageFillFixture,
  "image-crop.pptx": createImageCropFixture,
  "blip-effects.pptx": createBlipEffectsFixture,
  "image-stretch-tile.pptx": createImageStretchTileFixture,
  "metafile-images.pptx": createMetafileImagesFixture,
};
