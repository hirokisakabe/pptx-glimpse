/**
 * VRT (Visual Regression Testing) 用 PPTX フィクスチャ生成スクリプト
 *
 * 使い方: npx tsx scripts/vrt/create-fixtures.ts
 */
import JSZip from "jszip";
import sharp from "sharp";
import { writeFileSync, mkdirSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));

// --- Constants ---
const SLIDE_W = 9144000;
const SLIDE_H = 5143500;

const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
};

const REL_TYPES = {
  officeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
};

const COLORS = [
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47",
  "FF6384",
  "36A2EB",
  "FFCE56",
  "9966FF",
];

// --- Common XML Templates ---
const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.officeDocument}" Target="ppt/presentation.xml"/>
</Relationships>`;

const slideMaster1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

const slideMaster1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="${REL_TYPES.theme}" Target="../theme/theme1.xml"/>
</Relationships>`;

const slideLayout1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="blank">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const slideLayout1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideMaster}" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const theme1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${NS.a}" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

// --- Helper Functions ---

interface GridPos {
  x: number;
  y: number;
  w: number;
  h: number;
}

function gridPosition(
  col: number,
  row: number,
  cols: number,
  rows: number,
  margin = 200000,
): GridPos {
  const cellW = (SLIDE_W - margin * (cols + 1)) / cols;
  const cellH = (SLIDE_H - margin * (rows + 1)) / rows;
  return {
    x: margin + col * (cellW + margin),
    y: margin + row * (cellH + margin),
    w: cellW,
    h: cellH,
  };
}

function shapeXml(
  id: number,
  name: string,
  opts: {
    preset: string;
    x: number;
    y: number;
    cx: number;
    cy: number;
    fillXml?: string;
    outlineXml?: string;
    textBodyXml?: string;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    adjValues?: { name: string; val: number }[];
  },
): string {
  const rot = opts.rotation ? ` rot="${opts.rotation * 60000}"` : "";
  const fH = opts.flipH ? ` flipH="1"` : "";
  const fV = opts.flipV ? ` flipV="1"` : "";
  const adjList =
    opts.adjValues && opts.adjValues.length > 0
      ? opts.adjValues.map((a) => `<a:gd name="${a.name}" fmla="val ${a.val}"/>`).join("")
      : "";
  const fill =
    opts.fillXml ?? `<a:solidFill><a:srgbClr val="${COLORS[id % COLORS.length]}"/></a:solidFill>`;
  const outline = opts.outlineXml ?? "";
  const txBody = opts.textBodyXml ?? "";

  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm${rot}${fH}${fV}><a:off x="${opts.x}" y="${opts.y}"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>
    <a:prstGeom prst="${opts.preset}"><a:avLst>${adjList}</a:avLst></a:prstGeom>
    ${fill}
    ${outline}
  </p:spPr>
  ${txBody}
</p:sp>`;
}

function solidFillXml(color: string): string {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

function gradientFillXml(stops: { pos: number; color: string }[], angle: number): string {
  const gsItems = stops
    .map((s) => `<a:gs pos="${s.pos}"><a:srgbClr val="${s.color}"/></a:gs>`)
    .join("");
  return `<a:gradFill><a:gsLst>${gsItems}</a:gsLst><a:lin ang="${angle}" scaled="1"/></a:gradFill>`;
}

function outlineXml(width: number, color: string, dashStyle?: string): string {
  const dash = dashStyle ? `<a:prstDash val="${dashStyle}"/>` : "";
  return `<a:ln w="${width}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${dash}</a:ln>`;
}

function textBodyXmlHelper(
  text: string,
  opts?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    fontSize?: number;
    color?: string;
    align?: string;
    anchor?: string;
    wrap?: string;
    lineSpacing?: number;
    normAutofit?: { fontScale: number; lnSpcReduction?: number };
  },
): string {
  const sz = opts?.fontSize ? ` sz="${opts.fontSize * 100}"` : ` sz="1400"`;
  const b = opts?.bold ? ` b="1"` : "";
  const i = opts?.italic ? ` i="1"` : "";
  const u = opts?.underline ? ` u="sng"` : "";
  const strike = opts?.strikethrough ? ` strike="sngStrike"` : "";
  const fillColor = opts?.color ?? "000000";
  const algn = opts?.align ? ` algn="${opts.align}"` : "";
  const anchor = opts?.anchor ?? "ctr";
  const wrap = opts?.wrap ? ` wrap="${opts.wrap}"` : "";
  const lnSpc = opts?.lineSpacing ? `<a:lnSpc><a:spcPct val="${opts.lineSpacing}"/></a:lnSpc>` : "";
  let autofitXml = "";
  if (opts?.normAutofit) {
    const lnSpcR = opts.normAutofit.lnSpcReduction
      ? ` lnSpcReduction="${opts.normAutofit.lnSpcReduction}"`
      : "";
    autofitXml = `<a:normAutofit fontScale="${opts.normAutofit.fontScale}"${lnSpcR}/>`;
  }

  return `<p:txBody>
  <a:bodyPr anchor="${anchor}"${wrap}>${autofitXml}</a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr${algn}>${lnSpc}</a:pPr>
    <a:r>
      <a:rPr lang="en-US"${sz}${b}${i}${u}${strike}>
        <a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill>
      </a:rPr>
      <a:t>${text}</a:t>
    </a:r>
  </a:p>
</p:txBody>`;
}

function wrapSlideXml(spTreeContent: string, backgroundXml = ""): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    ${backgroundXml}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      ${spTreeContent}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function slideRelsXml(extraRels: { id: string; type: string; target: string }[] = []): string {
  const extras = extraRels
    .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`)
    .join("\n  ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
  ${extras}
</Relationships>`;
}

interface SlideData {
  xml: string;
  rels: string;
}

interface PptxBuildOptions {
  slides: SlideData[];
  charts?: Map<string, string>;
  media?: Map<string, Buffer>;
  contentTypesExtra?: string[];
  slideMasterXml?: string;
  slideMasterRelsXml?: string;
}

async function buildPptx(options: PptxBuildOptions): Promise<Buffer> {
  const zip = new JSZip();

  // Content_Types
  const slideOverrides = options.slides
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    )
    .join("\n  ");
  const extraOverrides = (options.contentTypesExtra ?? []).join("\n  ");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideOverrides}
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${extraOverrides}
</Types>`,
  );

  zip.file("_rels/.rels", rootRels);

  // Presentation
  const sldIdLst = options.slides
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${2 + i}"/>`)
    .join("");
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:sldMasterIdLst><p:sldMasterId r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${sldIdLst}</p:sldIdLst>
  <p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}" type="screen16x9"/>
</p:presentation>`,
  );

  // Presentation rels
  const slideRels = options.slides
    .map(
      (_, i) =>
        `<Relationship Id="rId${2 + i}" Type="${REL_TYPES.slide}" Target="slides/slide${i + 1}.xml"/>`,
    )
    .join("\n  ");
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideMaster}" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
  <Relationship Id="rId${2 + options.slides.length}" Type="${REL_TYPES.theme}" Target="theme/theme1.xml"/>
</Relationships>`,
  );

  // Slides
  options.slides.forEach((slide, i) => {
    zip.file(`ppt/slides/slide${i + 1}.xml`, slide.xml);
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, slide.rels);
  });

  // Common
  zip.file("ppt/slideMasters/slideMaster1.xml", options.slideMasterXml ?? slideMaster1);
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    options.slideMasterRelsXml ?? slideMaster1Rels,
  );
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", theme1);

  // Charts
  if (options.charts) {
    for (const [path, xml] of options.charts) {
      zip.file(path, xml);
    }
  }

  // Media
  if (options.media) {
    for (const [path, buf] of options.media) {
      zip.file(path, buf);
    }
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

const FIXTURE_OUT_DIR = join(__dirname, "../../tests/vrt/internal/fixtures");

function savePptx(buffer: Buffer, name: string): void {
  mkdirSync(FIXTURE_OUT_DIR, { recursive: true });
  const path = join(FIXTURE_OUT_DIR, name);
  writeFileSync(path, buffer);
  console.log(`  Created: ${path}`);
}

// ============================================================
// Fixture Generators
// ============================================================

// --- 1. Shapes ---
async function createShapesFixture(): Promise<void> {
  const presets1 = [
    "rect",
    "ellipse",
    "roundRect",
    "triangle",
    "rtTriangle",
    "diamond",
    "parallelogram",
    "trapezoid",
    "pentagon",
    "hexagon",
  ];
  const presets2 = [
    "star4",
    "star5",
    "rightArrow",
    "leftArrow",
    "upArrow",
    "downArrow",
    "line",
    "cloud",
    "heart",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      const lineXml = preset === "line" ? outlineXml(25400, "333333") : outlineXml(12700, "333333");
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: lineXml,
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(presets1, 5, 2);
  const slide2 = makeSlide(presets2, 5, 2);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-shapes.pptx");
}

// --- 2. Fill and Lines ---
async function createFillAndLinesFixture(): Promise<void> {
  // Slide 1: Fill types
  let id = 2;
  const fillShapes: string[] = [];

  // Solid fills
  const solidColors = ["FF6384", "36A2EB", "FFCE56"];
  solidColors.forEach((c, i) => {
    const pos = gridPosition(i, 0, 3, 2);
    fillShapes.push(
      shapeXml(id++, `solid-${c}`, {
        preset: "roundRect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(c),
      }),
    );
  });

  // Gradient fills
  const gradients = [
    { angle: 0, label: "horizontal" },
    { angle: 5400000, label: "vertical" },
    { angle: 2700000, label: "diagonal" },
  ];
  gradients.forEach((g, i) => {
    const pos = gridPosition(i, 1, 3, 2);
    fillShapes.push(
      shapeXml(id++, `grad-${g.label}`, {
        preset: "roundRect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: gradientFillXml(
          [
            { pos: 0, color: "4472C4" },
            { pos: 100000, color: "ED7D31" },
          ],
          g.angle,
        ),
      }),
    );
  });

  const slide1 = wrapSlideXml(fillShapes.join("\n"));

  // Slide 2: Line dash styles
  id = 2;
  const dashStyles = [
    "solid",
    "dash",
    "dot",
    "dashDot",
    "lgDash",
    "lgDashDot",
    "sysDash",
    "sysDot",
  ];
  const lineShapes = dashStyles.map((ds, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const pos = gridPosition(col, row, 4, 2);
    return shapeXml(id++, `line-${ds}`, {
      preset: "rect",
      x: pos.x,
      y: pos.y,
      cx: pos.w,
      cy: pos.h,
      fillXml: `<a:noFill/>`,
      outlineXml: outlineXml(25400, "333333", ds === "solid" ? undefined : ds),
      textBodyXml: textBodyXmlHelper(ds, { fontSize: 12, color: "333333" }),
    });
  });

  const slide2 = wrapSlideXml(lineShapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-fill-and-lines.pptx");
}

// --- 3. Text ---
async function createTextFixture(): Promise<void> {
  // Slide 1: Text formatting
  let id = 2;
  const textShapes1: string[] = [];
  const textTests = [
    { label: "Bold", bold: true },
    { label: "Italic", italic: true },
    { label: "Underline", underline: true },
    { label: "Strike", strikethrough: true },
    { label: "Small (12pt)", fontSize: 12 },
    { label: "Large (36pt)", fontSize: 36 },
    { label: "Red Text", color: "FF0000" },
    { label: "Blue Text", color: "0000FF" },
  ];
  textTests.forEach((t, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const pos = gridPosition(col, row, 4, 2);
    textShapes1.push(
      shapeXml(id++, `text-${t.label}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml("F0F0F0"),
        textBodyXml: textBodyXmlHelper(t.label, {
          bold: t.bold,
          italic: t.italic,
          underline: t.underline,
          strikethrough: t.strikethrough,
          fontSize: t.fontSize ?? 18,
          color: t.color ?? "333333",
        }),
      }),
    );
  });
  const slide1 = wrapSlideXml(textShapes1.join("\n"));

  // Slide 2: Alignment, line spacing, autofit, wrapping
  id = 2;
  const textShapes2: string[] = [];
  const alignTests = [
    { label: "Left Align", align: "l" },
    { label: "Center", align: "ctr" },
    { label: "Right Align", align: "r" },
    { label: "Top Anchor", align: "ctr", anchor: "t" },
    { label: "Line Spacing 200%", lineSpacing: 200000 },
    {
      label: "AutoFit Text That Should Shrink Down",
      normAutofit: { fontScale: 50000 },
    },
  ];
  alignTests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    textShapes2.push(
      shapeXml(id++, `text-${t.label}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml("F0F0F0"),
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 18,
          color: "333333",
          align: t.align,
          anchor: t.anchor,
          lineSpacing: t.lineSpacing,
          normAutofit: t.normAutofit,
        }),
      }),
    );
  });
  const slide2 = wrapSlideXml(textShapes2.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-text.pptx");
}

// --- 4. Transform ---
async function createTransformFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];
  const tests = [
    { label: "Rot 45", rotation: 45 },
    { label: "Rot 90", rotation: 90 },
    { label: "FlipH", flipH: true },
    { label: "FlipV", flipV: true },
    { label: "FlipH+V", flipH: true, flipV: true },
    { label: "Rot+FlipH", rotation: 30, flipH: true },
  ];
  tests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    shapes.push(
      shapeXml(id++, t.label, {
        preset: "rightArrow",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i]),
        outlineXml: outlineXml(12700, "333333"),
        rotation: t.rotation,
        flipH: t.flipH,
        flipV: t.flipV,
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 12,
          color: "FFFFFF",
        }),
      }),
    );
  });
  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vrt-transform.pptx");
}

// --- 5. Background ---
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
  savePptx(buffer, "vrt-background.pptx");
}

// --- 6. Groups ---
async function createGroupsFixture(): Promise<void> {
  // Group containing a rect, ellipse, and text shape
  const grpX = 500000;
  const grpY = 500000;
  const grpW = 8000000;
  const grpH = 4000000;

  const groupContent = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="2" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="${grpX}" y="${grpY}"/>
      <a:ext cx="${grpW}" cy="${grpH}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grpW}" cy="${grpH}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${shapeXml(3, "GroupRect", {
    preset: "rect",
    x: 0,
    y: 0,
    cx: 3500000,
    cy: 3500000,
    fillXml: solidFillXml("4472C4"),
    textBodyXml: textBodyXmlHelper("In Group", { fontSize: 18, color: "FFFFFF" }),
  })}
  ${shapeXml(4, "GroupEllipse", {
    preset: "ellipse",
    x: 4000000,
    y: 0,
    cx: 3500000,
    cy: 3500000,
    fillXml: solidFillXml("ED7D31"),
    textBodyXml: textBodyXmlHelper("Ellipse", { fontSize: 18, color: "FFFFFF" }),
  })}
</p:grpSp>`;

  const slide = wrapSlideXml(groupContent);
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vrt-groups.pptx");
}

// --- 7. Charts ---
function chartXml(
  chartType: string,
  opts: {
    barDir?: string;
    title?: string;
    legendPos?: string;
    series: {
      name: string;
      categories?: string[];
      values: number[];
      xValues?: number[];
    }[];
  },
): string {
  const titleXml = opts.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${opts.title}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
    : "";
  const legendXml = opts.legendPos
    ? `<c:legend><c:legendPos val="${opts.legendPos}"/></c:legend>`
    : "";

  const seriesXml = opts.series
    .map((s, i) => {
      const nameXml = `<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${s.name}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
      const catXml = s.categories
        ? `<c:cat><c:strRef><c:strCache>${s.categories.map((c, j) => `<c:pt idx="${j}"><c:v>${c}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>`
        : "";
      const valTag = chartType === "scatterChart" ? "c:yVal" : "c:val";
      const valXml = `<${valTag}><c:numRef><c:numCache>${s.values.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></${valTag}>`;
      const xValXml = s.xValues
        ? `<c:xVal><c:numRef><c:numCache>${s.xValues.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:xVal>`
        : "";
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nameXml}${catXml}${xValXml}${valXml}</c:ser>`;
    })
    .join("");

  const barDirXml = opts.barDir ? `<c:barDir val="${opts.barDir}"/>` : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}">
  <c:chart>
    ${titleXml}
    <c:plotArea>
      <c:${chartType}>
        ${barDirXml}
        ${seriesXml}
      </c:${chartType}>
    </c:plotArea>
    ${legendXml}
  </c:chart>
</c:chartSpace>`;
}

function graphicFrameXml(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  chartRId: string,
): string {
  return `<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart xmlns:c="${NS.c}" r:id="${chartRId}"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;
}

async function createChartsFixture(): Promise<void> {
  const charts = new Map<string, string>();
  const slides: SlideData[] = [];

  const margin = 300000;

  // Slide 1: Bar chart
  const barChart = chartXml("barChart", {
    barDir: "col",
    title: "Sales by Quarter",
    legendPos: "b",
    series: [
      { name: "FY2024", categories: ["Q1", "Q2", "Q3", "Q4"], values: [10, 25, 15, 30] },
      { name: "FY2025", categories: ["Q1", "Q2", "Q3", "Q4"], values: [15, 20, 25, 35] },
    ],
  });
  charts.set("ppt/charts/chart1.xml", barChart);
  const gf1 = graphicFrameXml(
    2,
    "Bar Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf1),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart1.xml" }]),
  });

  // Slide 2: Line chart
  const lineChart = chartXml("lineChart", {
    title: "Monthly Trend",
    legendPos: "b",
    series: [
      {
        name: "Revenue",
        categories: ["Jan", "Feb", "Mar", "Apr", "May"],
        values: [100, 120, 90, 150, 130],
      },
    ],
  });
  charts.set("ppt/charts/chart2.xml", lineChart);
  const gf2 = graphicFrameXml(
    2,
    "Line Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf2),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart2.xml" }]),
  });

  // Slide 3: Pie chart
  const pieChart = chartXml("pieChart", {
    title: "Market Share",
    legendPos: "r",
    series: [{ name: "Share", categories: ["A", "B", "C", "D"], values: [40, 25, 20, 15] }],
  });
  charts.set("ppt/charts/chart3.xml", pieChart);
  const gf3 = graphicFrameXml(
    2,
    "Pie Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf3),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart3.xml" }]),
  });

  // Slide 4: Scatter chart
  const scatterChart = chartXml("scatterChart", {
    title: "Data Points",
    legendPos: "b",
    series: [{ name: "Dataset", xValues: [1, 2, 3, 5, 8], values: [2, 4, 3, 7, 6] }],
  });
  charts.set("ppt/charts/chart4.xml", scatterChart);
  const gf4 = graphicFrameXml(
    2,
    "Scatter Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf4),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart4.xml" }]),
  });

  const buffer = await buildPptx({
    slides,
    charts,
    contentTypesExtra: [
      `<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart4.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    ],
  });
  savePptx(buffer, "vrt-charts.pptx");
}

// --- 8. Connectors ---
async function createConnectorsFixture(): Promise<void> {
  const connectors = [
    {
      id: 2,
      x: 500000,
      y: 500000,
      cx: 3000000,
      cy: 0,
      color: "4472C4",
      dash: undefined as string | undefined,
    },
    { id: 3, x: 500000, y: 1500000, cx: 3000000, cy: 2000000, color: "ED7D31", dash: "dash" },
    { id: 4, x: 5000000, y: 500000, cx: 0, cy: 4000000, color: "70AD47", dash: "dot" },
  ];

  const cxnXml = connectors
    .map((c) => {
      const dashXml = c.dash ? `<a:prstDash val="${c.dash}"/>` : "";
      return `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="${c.id}" name="Connector ${c.id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${c.x}" y="${c.y}"/><a:ext cx="${c.cx}" cy="${c.cy}"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="${c.color}"/></a:solidFill>${dashXml}</a:ln>
  </p:spPr>
</p:cxnSp>`;
    })
    .join("\n");

  const slide = wrapSlideXml(cxnXml);
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vrt-connectors.pptx");
}

// --- 9. Custom Geometry ---
async function createCustomGeometryFixture(): Promise<void> {
  // A custom star-like path and a custom curve
  const customShape1 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="CustomStar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3500000" cy="3500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:ahLst/>
      <a:cxnLst/>
      <a:rect l="0" t="0" r="0" b="0"/>
      <a:pathLst>
        <a:path w="1000" h="1000">
          <a:moveTo><a:pt x="500" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="650" y="350"/></a:lnTo>
          <a:lnTo><a:pt x="1000" y="400"/></a:lnTo>
          <a:lnTo><a:pt x="750" y="650"/></a:lnTo>
          <a:lnTo><a:pt x="800" y="1000"/></a:lnTo>
          <a:lnTo><a:pt x="500" y="850"/></a:lnTo>
          <a:lnTo><a:pt x="200" y="1000"/></a:lnTo>
          <a:lnTo><a:pt x="250" y="650"/></a:lnTo>
          <a:lnTo><a:pt x="0" y="400"/></a:lnTo>
          <a:lnTo><a:pt x="350" y="350"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="FFC000"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  const customShape2 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="CustomCurve"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="5000000" y="500000"/><a:ext cx="3500000" cy="3500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:ahLst/>
      <a:cxnLst/>
      <a:rect l="0" t="0" r="0" b="0"/>
      <a:pathLst>
        <a:path w="1000" h="1000">
          <a:moveTo><a:pt x="0" y="500"/></a:moveTo>
          <a:cubicBezTo>
            <a:pt x="250" y="0"/>
            <a:pt x="750" y="1000"/>
            <a:pt x="1000" y="500"/>
          </a:cubicBezTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  const slide = wrapSlideXml(customShape1 + "\n" + customShape2);
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vrt-custom-geometry.pptx");
}

// --- 10. Image ---
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
  savePptx(buffer, "vrt-image.pptx");
}

// --- 11. Tables ---
function tableGraphicFrameXml(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  tblXml: string,
): string {
  return `<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      ${tblXml}
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;
}

function tableCellXml(
  text: string,
  opts?: {
    fillColor?: string;
    bold?: boolean;
    fontSize?: number;
    fontColor?: string;
    borderColor?: string;
    borderWidth?: number;
    gridSpan?: number;
    rowSpan?: number;
    hMerge?: boolean;
    vMerge?: boolean;
  },
): string {
  const spanAttrs = [
    opts?.gridSpan && opts.gridSpan > 1 ? ` gridSpan="${opts.gridSpan}"` : "",
    opts?.rowSpan && opts.rowSpan > 1 ? ` rowSpan="${opts.rowSpan}"` : "",
  ].join("");

  const sz = opts?.fontSize ? ` sz="${opts.fontSize * 100}"` : ` sz="1200"`;
  const b = opts?.bold ? ` b="1"` : "";
  const fontColor = opts?.fontColor ?? "000000";

  const fillXml = opts?.fillColor
    ? `<a:solidFill><a:srgbClr val="${opts.fillColor}"/></a:solidFill>`
    : "";

  const bw = opts?.borderWidth ?? 12700;
  const bc = opts?.borderColor ?? "000000";
  const borderXml = `<a:lnL w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnL>
      <a:lnR w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnR>
      <a:lnT w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnT>
      <a:lnB w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnB>`;

  const hMergeAttr = opts?.hMerge ? ` hMerge="1"` : "";
  const vMergeAttr = opts?.vMerge ? ` vMerge="1"` : "";

  return `<a:tc${spanAttrs}>
    <a:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US"${sz}${b}>
            <a:solidFill><a:srgbClr val="${fontColor}"/></a:solidFill>
          </a:rPr>
          <a:t>${text}</a:t>
        </a:r>
      </a:p>
    </a:txBody>
    <a:tcPr${hMergeAttr}${vMergeAttr}>
      ${borderXml}
      ${fillXml}
    </a:tcPr>
  </a:tc>`;
}

async function createTablesFixture(): Promise<void> {
  const margin = 300000;

  // Slide 1: Basic table with header row
  const colW = 2286000; // 3 columns
  const rowH = 457200;
  const tblW = colW * 3;
  const tblH = rowH * 4;

  const headerRow = `<a:tr h="${rowH}">
    ${tableCellXml("Name", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
    ${tableCellXml("Value", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
    ${tableCellXml("Status", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
  </a:tr>`;

  const dataRows = [
    ["Alpha", "100", "Active"],
    ["Beta", "250", "Pending"],
    ["Gamma", "75", "Done"],
  ]
    .map(
      (row, i) =>
        `<a:tr h="${rowH}">
    ${row.map((cell) => tableCellXml(cell, { fillColor: i % 2 === 0 ? "D6E4F0" : "FFFFFF" })).join("\n    ")}
  </a:tr>`,
    )
    .join("\n  ");

  const tbl1 = `<a:tbl>
    <a:tblPr firstRow="1"/>
    <a:tblGrid>
      <a:gridCol w="${colW}"/>
      <a:gridCol w="${colW}"/>
      <a:gridCol w="${colW}"/>
    </a:tblGrid>
    ${headerRow}
    ${dataRows}
  </a:tbl>`;

  const gf1 = tableGraphicFrameXml(2, "Basic Table", margin, margin, tblW, tblH, tbl1);
  const slide1 = wrapSlideXml(gf1);

  // Slide 2: Table with cell merging
  const col2W = 2286000;
  const row2H = 600000;
  const tbl2W = col2W * 3;
  const tbl2H = row2H * 3;

  const tbl2 = `<a:tbl>
    <a:tblPr/>
    <a:tblGrid>
      <a:gridCol w="${col2W}"/>
      <a:gridCol w="${col2W}"/>
      <a:gridCol w="${col2W}"/>
    </a:tblGrid>
    <a:tr h="${row2H}">
      ${tableCellXml("Merged Header (3 cols)", { fillColor: "ED7D31", fontColor: "FFFFFF", bold: true, gridSpan: 3 })}
      ${tableCellXml("", { hMerge: true })}
      ${tableCellXml("", { hMerge: true })}
    </a:tr>
    <a:tr h="${row2H}">
      ${tableCellXml("Merged\nRows", { fillColor: "A5A5A5", fontColor: "FFFFFF", rowSpan: 2 })}
      ${tableCellXml("Cell B2", { fillColor: "FFFFFF" })}
      ${tableCellXml("Cell C2", { fillColor: "FFFFFF" })}
    </a:tr>
    <a:tr h="${row2H}">
      ${tableCellXml("", { vMerge: true })}
      ${tableCellXml("Cell B3", { fillColor: "D6E4F0" })}
      ${tableCellXml("Cell C3", { fillColor: "D6E4F0" })}
    </a:tr>
  </a:tbl>`;

  const gf2 = tableGraphicFrameXml(2, "Merge Table", margin, margin, tbl2W, tbl2H, tbl2);
  const slide2 = wrapSlideXml(gf2);

  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-tables.pptx");
}

// --- 12. Bullets and Numbering ---
function bulletParagraphsXml(
  items: {
    text: string;
    bulletXml: string;
    marL?: number;
    indent?: number;
    lvl?: number;
  }[],
  opts?: { anchor?: string },
): string {
  const anchor = opts?.anchor ?? "t";
  const paragraphs = items
    .map((item) => {
      const marL = item.marL ?? 342900;
      const indent = item.indent ?? -342900;
      const lvl = item.lvl ?? 0;
      return `<a:p>
      <a:pPr lvl="${lvl}" marL="${marL}" indent="${indent}">
        ${item.bulletXml}
      </a:pPr>
      <a:r>
        <a:rPr lang="en-US" sz="1400">
          <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
        </a:rPr>
        <a:t>${item.text}</a:t>
      </a:r>
    </a:p>`;
    })
    .join("\n");

  return `<p:txBody>
  <a:bodyPr anchor="${anchor}"/>
  <a:lstStyle/>
  ${paragraphs}
</p:txBody>`;
}

async function createBulletsFixture(): Promise<void> {
  // Slide 1: Bullet characters (buChar)
  let id = 2;
  const shapes1: string[] = [];

  // buChar - standard bullet
  const buCharItems = [
    { text: "First bullet item", bulletXml: `<a:buChar char="\u2022"/>` },
    { text: "Second bullet item", bulletXml: `<a:buChar char="\u2022"/>` },
    { text: "Third bullet item", bulletXml: `<a:buChar char="\u2022"/>` },
  ];
  const pos1 = gridPosition(0, 0, 2, 2);
  shapes1.push(
    shapeXml(id++, "buChar-standard", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buCharItems),
    }),
  );

  // buChar - dash bullet
  const buDashItems = [
    { text: "Dash item A", bulletXml: `<a:buChar char="-"/>` },
    { text: "Dash item B", bulletXml: `<a:buChar char="-"/>` },
  ];
  const pos2 = gridPosition(1, 0, 2, 2);
  shapes1.push(
    shapeXml(id++, "buChar-dash", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buDashItems),
    }),
  );

  // buAutoNum - arabicPeriod
  const buNumItems = [
    {
      text: "Numbered one",
      bulletXml: `<a:buAutoNum type="arabicPeriod"/>`,
    },
    {
      text: "Numbered two",
      bulletXml: `<a:buAutoNum type="arabicPeriod"/>`,
    },
    {
      text: "Numbered three",
      bulletXml: `<a:buAutoNum type="arabicPeriod"/>`,
    },
  ];
  const pos3 = gridPosition(0, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "buAutoNum-arabic", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buNumItems),
    }),
  );

  // buAutoNum - alphaLcPeriod
  const buAlphaItems = [
    {
      text: "Alpha item",
      bulletXml: `<a:buAutoNum type="alphaLcPeriod"/>`,
    },
    {
      text: "Beta item",
      bulletXml: `<a:buAutoNum type="alphaLcPeriod"/>`,
    },
    {
      text: "Gamma item",
      bulletXml: `<a:buAutoNum type="alphaLcPeriod"/>`,
    },
  ];
  const pos4 = gridPosition(1, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "buAutoNum-alpha", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buAlphaItems),
    }),
  );

  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: buNone, buFont, mixed
  id = 2;
  const shapes2: string[] = [];

  // buNone
  const buNoneItems = [
    { text: "No bullet here", bulletXml: `<a:buNone/>` },
    { text: "Also no bullet", bulletXml: `<a:buNone/>` },
  ];
  const pos5 = gridPosition(0, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "buNone", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buNoneItems),
    }),
  );

  // buFont + buChar (custom font bullet)
  const buFontItems = [
    {
      text: "Custom font bullet",
      bulletXml: `<a:buFont typeface="Arial"/><a:buChar char="\u25A0"/>`,
    },
    {
      text: "Another custom",
      bulletXml: `<a:buFont typeface="Arial"/><a:buChar char="\u25A0"/>`,
    },
  ];
  const pos6 = gridPosition(1, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "buFont-custom", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buFontItems),
    }),
  );

  // romanUcPeriod numbering
  const buRomanItems = [
    {
      text: "Roman one",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
    {
      text: "Roman two",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
    {
      text: "Roman three",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
    {
      text: "Roman four",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
  ];
  const pos7 = gridPosition(0, 1, 2, 2);
  shapes2.push(
    shapeXml(id++, "buAutoNum-roman", {
      preset: "rect",
      x: pos7.x,
      y: pos7.y,
      cx: pos7.w,
      cy: pos7.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buRomanItems),
    }),
  );

  // Colored bullet
  const buColorItems = [
    {
      text: "Red bullet",
      bulletXml: `<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buChar char="\u2022"/>`,
    },
    {
      text: "Blue bullet",
      bulletXml: `<a:buClr><a:srgbClr val="0000FF"/></a:buClr><a:buChar char="\u2022"/>`,
    },
    {
      text: "Green bullet",
      bulletXml: `<a:buClr><a:srgbClr val="00AA00"/></a:buClr><a:buChar char="\u2022"/>`,
    },
  ];
  const pos8 = gridPosition(1, 1, 2, 2);
  shapes2.push(
    shapeXml(id++, "buClr-colored", {
      preset: "rect",
      x: pos8.x,
      y: pos8.y,
      cx: pos8.w,
      cy: pos8.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buColorItems),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));

  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-bullets.pptx");
}

// --- 13. Flowchart Shapes ---
async function createFlowchartFixture(): Promise<void> {
  const flowchartPresets1 = [
    "flowChartProcess",
    "flowChartAlternateProcess",
    "flowChartDecision",
    "flowChartInputOutput",
    "flowChartPredefinedProcess",
    "flowChartInternalStorage",
    "flowChartDocument",
    "flowChartMultidocument",
    "flowChartTerminator",
    "flowChartPreparation",
    "flowChartManualInput",
    "flowChartManualOperation",
    "flowChartConnector",
    "flowChartOffpageConnector",
  ];
  const flowchartPresets2 = [
    "flowChartPunchedCard",
    "flowChartPunchedTape",
    "flowChartCollate",
    "flowChartSort",
    "flowChartExtract",
    "flowChartMerge",
    "flowChartOnlineStorage",
    "flowChartDelay",
    "flowChartDisplay",
    "flowChartMagneticTape",
    "flowChartMagneticDisk",
    "flowChartMagneticDrum",
    "flowChartSummingJunction",
    "flowChartOr",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(flowchartPresets1, 7, 2);
  const slide2 = makeSlide(flowchartPresets2, 7, 2);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-flowchart.pptx");
}

// --- 14. Callout and Arc Shapes ---
async function createCalloutsArcsFixture(): Promise<void> {
  const presets = [
    "wedgeRectCallout",
    "wedgeRoundRectCallout",
    "wedgeEllipseCallout",
    "cloudCallout",
    "borderCallout1",
    "borderCallout2",
    "borderCallout3",
    "arc",
    "chord",
    "pie",
    "blockArc",
  ];

  let id = 2;
  const shapes = presets.map((preset, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const pos = gridPosition(col, row, 4, 3);
    return shapeXml(id++, preset, {
      preset,
      x: pos.x,
      y: pos.y,
      cx: pos.w,
      cy: pos.h,
      fillXml: solidFillXml(COLORS[i % COLORS.length]),
      outlineXml: outlineXml(12700, "333333"),
    });
  });

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vrt-callouts-arcs.pptx");
}

// --- 15. Extended Arrows and Stars ---
async function createArrowsStarsFixture(): Promise<void> {
  const arrowPresets = [
    "leftRightArrow",
    "upDownArrow",
    "notchedRightArrow",
    "stripedRightArrow",
    "chevron",
    "homePlate",
    "bentArrow",
    "bendUpArrow",
    "quadArrow",
    "leftUpArrow",
  ];
  const starPresets = [
    "star6",
    "star8",
    "star10",
    "star12",
    "star16",
    "star24",
    "star32",
    "irregularSeal1",
    "irregularSeal2",
    "sun",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(arrowPresets, 5, 2);
  const slide2 = makeSlide(starPresets, 5, 2);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-arrows-stars.pptx");
}

// --- 16. Math and Other Shapes ---
async function createMathOtherFixture(): Promise<void> {
  const mathPresets = [
    "mathPlus",
    "mathMinus",
    "mathMultiply",
    "mathDivide",
    "mathEqual",
    "mathNotEqual",
  ];
  const otherPresets1 = [
    "plus",
    "heptagon",
    "octagon",
    "decagon",
    "dodecagon",
    "plaque",
    "can",
    "cube",
    "donut",
    "noSmoking",
    "smileyFace",
    "foldedCorner",
    "frame",
    "bevel",
  ];
  const otherPresets2 = [
    "halfFrame",
    "corner",
    "diagStripe",
    "snip1Rect",
    "snip2SameRect",
    "snip2DiagRect",
    "snipRoundRect",
    "round1Rect",
    "round2SameRect",
    "round2DiagRect",
    "lightningBolt",
    "moon",
    "teardrop",
    "wave",
    "doubleWave",
    "ribbon",
    "ribbon2",
    "bracketPair",
    "bracePair",
    "leftBracket",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(mathPresets, 3, 2);
  const slide2 = makeSlide(otherPresets1, 7, 2);
  const slide3 = makeSlide(otherPresets2, 7, 3);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
    ],
  });
  savePptx(buffer, "vrt-math-other.pptx");
}

// --- 17. Word Wrap ---
function multiRunTextBodyXml(
  paragraphs: {
    runs: {
      text: string;
      fontSize?: number;
      bold?: boolean;
      color?: string;
      lang?: string;
      baseline?: number;
    }[];
    align?: string;
  }[],
  opts?: { anchor?: string; wrap?: string },
): string {
  const anchor = opts?.anchor ?? "t";
  const wrap = opts?.wrap ? ` wrap="${opts.wrap}"` : "";
  const parasXml = paragraphs
    .map((para) => {
      const algn = para.align ? ` algn="${para.align}"` : "";
      const runsXml = para.runs
        .map((r) => {
          const sz = r.fontSize ? ` sz="${r.fontSize * 100}"` : ` sz="1400"`;
          const b = r.bold ? ` b="1"` : "";
          const lang = r.lang ?? "en-US";
          const fillColor = r.color ?? "000000";
          const bl = r.baseline !== undefined ? ` baseline="${r.baseline}"` : "";
          return `<a:r>
        <a:rPr lang="${lang}"${sz}${b}${bl}>
          <a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill>
        </a:rPr>
        <a:t>${r.text}</a:t>
      </a:r>`;
        })
        .join("\n    ");
      return `<a:p>
    <a:pPr${algn}/>
    ${runsXml}
  </a:p>`;
    })
    .join("\n  ");

  return `<p:txBody>
  <a:bodyPr anchor="${anchor}"${wrap}/>
  <a:lstStyle/>
  ${parasXml}
</p:txBody>`;
}

async function createWordWrapFixture(): Promise<void> {
  // Slide 1: Basic word wrap scenarios
  let id = 2;
  const shapes1: string[] = [];

  // 1. Long English text in normal-width shape
  const longEnText =
    "The quick brown fox jumps over the lazy dog. This is a long sentence that should wrap across multiple lines within the shape boundary.";
  const pos1 = gridPosition(0, 0, 2, 2);
  shapes1.push(
    shapeXml(id++, "long-en-text", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml([{ runs: [{ text: longEnText, fontSize: 14 }] }], {
        anchor: "t",
      }),
    }),
  );

  // 2. Long text in narrow shape
  const pos2 = { x: pos1.x + pos1.w + 200000, y: pos1.y, w: 1500000, h: pos1.h };
  shapes1.push(
    shapeXml(id++, "narrow-shape", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("E8F4FD"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [{ runs: [{ text: "Narrow shape forces frequent word wrapping here.", fontSize: 14 }] }],
        { anchor: "t" },
      ),
    }),
  );

  // 3. wrap="none" (no wrapping)
  const pos3 = gridPosition(0, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "no-wrap", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("FFF3E0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "This text has wrap=none so it should not wrap at the shape boundary.",
                fontSize: 14,
              },
            ],
          },
        ],
        { anchor: "t", wrap: "none" },
      ),
    }),
  );

  // 4. Japanese text wrapping
  const pos4 = gridPosition(1, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "japanese-text", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("F3E5F5"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "日本語のテキストは文字単位で折り返されます。長い文章を図形の中に配置した場合の表示を確認します。",
                fontSize: 14,
                lang: "ja-JP",
              },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: Advanced word wrap scenarios
  id = 2;
  const shapes2: string[] = [];

  // 1. Mixed font sizes in a single paragraph
  const pos5 = gridPosition(0, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "mixed-font-sizes", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("E8F5E9"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              { text: "Large ", fontSize: 28, bold: true, color: "1565C0" },
              { text: "and small ", fontSize: 12, color: "333333" },
              { text: "mixed ", fontSize: 20, color: "C62828" },
              { text: "in one paragraph that wraps across lines.", fontSize: 14, color: "333333" },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 2. Multiple paragraphs
  const pos6 = gridPosition(1, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "multi-paragraph", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("FFF8E1"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          { runs: [{ text: "First paragraph with enough text to wrap.", fontSize: 14 }] },
          {
            runs: [{ text: "Second paragraph also wraps within the shape.", fontSize: 14 }],
            align: "ctr",
          },
          {
            runs: [{ text: "Third paragraph right-aligned.", fontSize: 14 }],
            align: "r",
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 3. Text overflow (long text in small shape)
  const pos7 = {
    x: gridPosition(0, 1, 2, 2).x,
    y: gridPosition(0, 1, 2, 2).y,
    w: 2000000,
    h: 800000,
  };
  shapes2.push(
    shapeXml(id++, "text-overflow", {
      preset: "rect",
      x: pos7.x,
      y: pos7.y,
      cx: pos7.w,
      cy: pos7.h,
      fillXml: solidFillXml("FFEBEE"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "This text is too long for the small shape and will overflow beyond the visible area of the shape boundary.",
                fontSize: 14,
              },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 4. Mixed CJK and Latin text
  const pos8 = gridPosition(1, 1, 2, 2);
  shapes2.push(
    shapeXml(id++, "mixed-cjk-latin", {
      preset: "rect",
      x: pos8.x,
      y: pos8.y,
      cx: pos8.w,
      cy: pos8.h,
      fillXml: solidFillXml("E0F2F1"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "English and 日本語 mixed text. テキストの折り返しが正しく動作するか確認します。Word wrap test.",
                fontSize: 14,
              },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "vrt-word-wrap.pptx");
}

// --- 18. Background blipFill ---
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
  savePptx(buffer, "vrt-background-blipfill.pptx");
}

// --- 19. Composite (Shape + Text + Fill + Transform) ---
async function createCompositeFixture(): Promise<void> {
  // Slide 1: Non-rectangular shapes with text and different anchors
  let id = 2;
  const shapes1: string[] = [];

  const shapeTextTests = [
    { preset: "ellipse", label: "Ellipse Top", anchor: "t", fill: "4472C4" },
    { preset: "ellipse", label: "Ellipse Center", anchor: "ctr", fill: "5B9BD5" },
    { preset: "ellipse", label: "Ellipse Bottom", anchor: "b", fill: "2E75B6" },
    { preset: "diamond", label: "Diamond\nText", anchor: "ctr", fill: "ED7D31" },
    { preset: "triangle", label: "Tri", anchor: "ctr", fill: "70AD47" },
    {
      preset: "roundRect",
      label: "Multi-line\nRounded Rect\nwith Text",
      anchor: "ctr",
      fill: "FFC000",
    },
  ];
  shapeTextTests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    shapes1.push(
      shapeXml(id++, t.label, {
        preset: t.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(t.fill),
        outlineXml: outlineXml(12700, "333333"),
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 14,
          color: "FFFFFF",
          anchor: t.anchor,
          align: "ctr",
        }),
      }),
    );
  });
  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: Shape + Fill + Text + Outline combinations
  id = 2;
  const shapes2: string[] = [];

  // Gradient fill + text
  const pos2a = gridPosition(0, 0, 3, 2);
  shapes2.push(
    shapeXml(id++, "grad-text", {
      preset: "roundRect",
      x: pos2a.x,
      y: pos2a.y,
      cx: pos2a.w,
      cy: pos2a.h,
      fillXml: gradientFillXml(
        [
          { pos: 0, color: "1A1A2E" },
          { pos: 100000, color: "E94560" },
        ],
        5400000,
      ),
      outlineXml: outlineXml(12700, "333333"),
      textBodyXml: textBodyXmlHelper("Gradient Fill", {
        fontSize: 18,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  );

  // Thick outline + text
  const pos2b = gridPosition(1, 0, 3, 2);
  shapes2.push(
    shapeXml(id++, "thick-outline-text", {
      preset: "rect",
      x: pos2b.x,
      y: pos2b.y,
      cx: pos2b.w,
      cy: pos2b.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(76200, "4472C4"),
      textBodyXml: textBodyXmlHelper("Thick Outline", {
        fontSize: 16,
        color: "333333",
      }),
    }),
  );

  // Semi-transparent fill + text
  const pos2c = gridPosition(2, 0, 3, 2);
  shapes2.push(
    shapeXml(id++, "semi-transparent-text", {
      preset: "roundRect",
      x: pos2c.x,
      y: pos2c.y,
      cx: pos2c.w,
      cy: pos2c.h,
      fillXml: `<a:solidFill><a:srgbClr val="4472C4"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(25400, "333333"),
      textBodyXml: textBodyXmlHelper("50% Alpha Fill", {
        fontSize: 16,
        bold: true,
        color: "000000",
      }),
    }),
  );

  // Gradient fill ellipse + italic text
  const pos2d = gridPosition(0, 1, 3, 2);
  shapes2.push(
    shapeXml(id++, "grad-ellipse-text", {
      preset: "ellipse",
      x: pos2d.x,
      y: pos2d.y,
      cx: pos2d.w,
      cy: pos2d.h,
      fillXml: gradientFillXml(
        [
          { pos: 0, color: "FF6384" },
          { pos: 50000, color: "FFCE56" },
          { pos: 100000, color: "36A2EB" },
        ],
        2700000,
      ),
      outlineXml: outlineXml(19050, "666666"),
      textBodyXml: textBodyXmlHelper("Rainbow Ellipse", {
        fontSize: 14,
        italic: true,
        color: "FFFFFF",
      }),
    }),
  );

  // Dashed outline + bold text
  const pos2e = gridPosition(1, 1, 3, 2);
  shapes2.push(
    shapeXml(id++, "dashed-bold", {
      preset: "diamond",
      x: pos2e.x,
      y: pos2e.y,
      cx: pos2e.w,
      cy: pos2e.h,
      fillXml: solidFillXml("E7E6E6"),
      outlineXml: outlineXml(38100, "ED7D31", "dash"),
      textBodyXml: textBodyXmlHelper("Dashed", {
        fontSize: 14,
        bold: true,
        color: "ED7D31",
      }),
    }),
  );

  // Solid fill hexagon + underline text
  const pos2f = gridPosition(2, 1, 3, 2);
  shapes2.push(
    shapeXml(id++, "hex-underline", {
      preset: "hexagon",
      x: pos2f.x,
      y: pos2f.y,
      cx: pos2f.w,
      cy: pos2f.h,
      fillXml: solidFillXml("70AD47"),
      outlineXml: outlineXml(25400, "333333"),
      textBodyXml: textBodyXmlHelper("Hexagon", {
        fontSize: 14,
        underline: true,
        color: "FFFFFF",
      }),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));

  // Slide 3: Transform + Text combinations
  id = 2;
  const shapes3: string[] = [];

  const transformTextTests = [
    { label: "Rot45 + Text", preset: "rect", rotation: 45, fill: "4472C4" },
    {
      label: "Rot30 + Grad",
      preset: "roundRect",
      rotation: 30,
      fill: undefined as string | undefined,
      useGrad: true,
    },
    { label: "FlipH + Text", preset: "rightArrow", flipH: true, fill: "ED7D31" },
    { label: "FlipV + Text", preset: "triangle", flipV: true, fill: "70AD47" },
    { label: "Rot+FlipH", preset: "parallelogram", rotation: 15, flipH: true, fill: "FFC000" },
    { label: "Rot+FlipV", preset: "pentagon", rotation: 60, flipV: true, fill: "9966FF" },
  ];
  transformTextTests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    const fill = t.useGrad
      ? gradientFillXml(
          [
            { pos: 0, color: "667EEA" },
            { pos: 100000, color: "764BA2" },
          ],
          5400000,
        )
      : solidFillXml(t.fill!);
    shapes3.push(
      shapeXml(id++, t.label, {
        preset: t.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: fill,
        outlineXml: outlineXml(12700, "333333"),
        rotation: t.rotation,
        flipH: t.flipH,
        flipV: t.flipV,
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 12,
          bold: true,
          color: "FFFFFF",
        }),
      }),
    );
  });
  const slide3 = wrapSlideXml(shapes3.join("\n"));

  // Slide 4: Overlapping shapes (z-order and semi-transparency)
  id = 2;
  const shapes4: string[] = [];

  // Layer 1: large background rectangle
  shapes4.push(
    shapeXml(id++, "bg-rect", {
      preset: "rect",
      x: 500000,
      y: 300000,
      cx: 4000000,
      cy: 3000000,
      fillXml: solidFillXml("4472C4"),
      outlineXml: outlineXml(12700, "333333"),
      textBodyXml: textBodyXmlHelper("Back (Z=1)", {
        fontSize: 14,
        color: "FFFFFF",
        anchor: "t",
        align: "l",
      }),
    }),
  );

  // Layer 2: overlapping ellipse
  shapes4.push(
    shapeXml(id++, "mid-ellipse", {
      preset: "ellipse",
      x: 1500000,
      y: 1000000,
      cx: 3500000,
      cy: 2500000,
      fillXml: solidFillXml("ED7D31"),
      outlineXml: outlineXml(12700, "333333"),
      textBodyXml: textBodyXmlHelper("Middle (Z=2)", {
        fontSize: 14,
        color: "FFFFFF",
      }),
    }),
  );

  // Layer 3: small foreground rounded rect
  shapes4.push(
    shapeXml(id++, "front-roundRect", {
      preset: "roundRect",
      x: 2500000,
      y: 1500000,
      cx: 2500000,
      cy: 1800000,
      fillXml: solidFillXml("70AD47"),
      outlineXml: outlineXml(19050, "333333"),
      textBodyXml: textBodyXmlHelper("Front (Z=3)", {
        fontSize: 16,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  );

  // Semi-transparent overlapping group (right side)
  shapes4.push(
    shapeXml(id++, "alpha-rect1", {
      preset: "rect",
      x: 5500000,
      y: 500000,
      cx: 2500000,
      cy: 2500000,
      fillXml: `<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(12700, "990000"),
      textBodyXml: textBodyXmlHelper("Red 50%", {
        fontSize: 12,
        color: "000000",
        anchor: "t",
      }),
    }),
  );

  shapes4.push(
    shapeXml(id++, "alpha-rect2", {
      preset: "rect",
      x: 6200000,
      y: 1200000,
      cx: 2500000,
      cy: 2500000,
      fillXml: `<a:solidFill><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(12700, "000099"),
      textBodyXml: textBodyXmlHelper("Blue 50%", {
        fontSize: 12,
        color: "000000",
        anchor: "b",
      }),
    }),
  );

  shapes4.push(
    shapeXml(id++, "alpha-ellipse", {
      preset: "ellipse",
      x: 5800000,
      y: 2000000,
      cx: 2200000,
      cy: 2200000,
      fillXml: `<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="40000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(12700, "006600"),
      textBodyXml: textBodyXmlHelper("Green 40%", {
        fontSize: 12,
        color: "000000",
      }),
    }),
  );

  const slide4 = wrapSlideXml(shapes4.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
      { xml: slide4, rels },
    ],
  });
  savePptx(buffer, "vrt-composite.pptx");
}

// --- 20. Text Decoration (superscript / subscript) ---
async function createTextDecorationFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  const testCases = [
    {
      label: "Superscript",
      paragraphs: [
        {
          runs: [
            { text: "E = mc", fontSize: 20 },
            { text: "2", fontSize: 14, baseline: 30000 },
          ],
        },
      ],
    },
    {
      label: "Subscript",
      paragraphs: [
        {
          runs: [
            { text: "H", fontSize: 20 },
            { text: "2", fontSize: 14, baseline: -25000 },
            { text: "O", fontSize: 20 },
          ],
        },
      ],
    },
    {
      label: "Mixed",
      paragraphs: [
        {
          runs: [
            { text: "x", fontSize: 18 },
            { text: "n", fontSize: 12, baseline: 30000 },
            { text: " + y", fontSize: 18 },
            { text: "m", fontSize: 12, baseline: -25000 },
          ],
        },
      ],
    },
    {
      label: "Multi-line",
      paragraphs: [
        {
          runs: [
            { text: "Line 1 with ", fontSize: 16 },
            { text: "super", fontSize: 12, baseline: 30000 },
          ],
        },
        {
          runs: [
            { text: "Line 2 with ", fontSize: 16 },
            { text: "sub", fontSize: 12, baseline: -25000 },
          ],
        },
      ],
    },
  ];

  testCases.forEach((tc, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const pos = gridPosition(col, row, 2, 2);
    shapes.push(
      shapeXml(id++, tc.label, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml("F0F0F0"),
        outlineXml: outlineXml(12700, "CCCCCC"),
        textBodyXml: multiRunTextBodyXml(tc.paragraphs, { anchor: "ctr" }),
      }),
    );
  });

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vrt-text-decoration.pptx");
}

// --- Main ---
async function main(): Promise<void> {
  console.log("Creating VRT fixtures...\n");

  await createShapesFixture();
  await createFillAndLinesFixture();
  await createTextFixture();
  await createTransformFixture();
  await createBackgroundFixture();
  await createGroupsFixture();
  await createChartsFixture();
  await createConnectorsFixture();
  await createCustomGeometryFixture();
  await createImageFixture();
  await createTablesFixture();
  await createBulletsFixture();
  await createFlowchartFixture();
  await createCalloutsArcsFixture();
  await createArrowsStarsFixture();
  await createMathOtherFixture();
  await createWordWrapFixture();
  await createBackgroundBlipFillFixture();
  await createCompositeFixture();
  await createTextDecorationFixture();

  console.log("\nDone!");
}

main().catch(console.error);
