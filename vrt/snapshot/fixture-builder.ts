/**
 * Shared PPTX fixture builder utilities for snapshot VRT.
 */
import { mkdirSync, writeFileSync } from "fs";
import JSZip from "jszip";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

export type FixtureCreator = () => Promise<void>;
export type FixtureCreatorMap = Record<string, FixtureCreator>;

// --- Constants ---
export const SLIDE_W = 9144000;
export const SLIDE_H = 5143500;

// 4:3 slide size
export const SLIDE_W_4_3 = 9144000;
export const SLIDE_H_4_3 = 6858000;

export const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
};

export const REL_TYPES = {
  officeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
};

export const COLORS = [
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

export interface GridPos {
  x: number;
  y: number;
  w: number;
  h: number;
}

export function gridPosition(
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

export function shapeXml(
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
    effectsXml?: string;
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
  const effects = opts.effectsXml ? `<a:effectLst>${opts.effectsXml}</a:effectLst>` : "";

  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm${rot}${fH}${fV}><a:off x="${opts.x}" y="${opts.y}"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>
    <a:prstGeom prst="${opts.preset}"><a:avLst>${adjList}</a:avLst></a:prstGeom>
    ${fill}
    ${outline}
    ${effects}
  </p:spPr>
  ${txBody}
</p:sp>`;
}

export function solidFillXml(color: string): string {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

export function gradientFillXml(stops: { pos: number; color: string }[], angle: number): string {
  const gsItems = stops
    .map((s) => `<a:gs pos="${s.pos}"><a:srgbClr val="${s.color}"/></a:gs>`)
    .join("");
  return `<a:gradFill><a:gsLst>${gsItems}</a:gsLst><a:lin ang="${angle}" scaled="1"/></a:gradFill>`;
}

export function outlineXml(
  width: number,
  color: string,
  dashStyle?: string,
  opts?: { cap?: string; join?: string },
): string {
  const capAttr = opts?.cap ? ` cap="${opts.cap}"` : "";
  const dash = dashStyle ? `<a:prstDash val="${dashStyle}"/>` : "";
  const joinXml = opts?.join ? `<a:${opts.join}/>` : "";
  return `<a:ln w="${width}"${capAttr}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${dash}${joinXml}</a:ln>`;
}

export function textBodyXmlHelper(
  text: string,
  opts?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    fontSize?: number;
    color?: string;
    typeface?: string;
    align?: string;
    anchor?: string;
    wrap?: string;
    lineSpacing?: number;
    lineSpacingPts?: number;
    normAutofit?: { fontScale: number; lnSpcReduction?: number };
  },
): string {
  const sz = opts?.fontSize ? ` sz="${opts.fontSize * 100}"` : ` sz="1400"`;
  const b = opts?.bold ? ` b="1"` : "";
  const i = opts?.italic ? ` i="1"` : "";
  const u = opts?.underline ? ` u="sng"` : "";
  const strike = opts?.strikethrough ? ` strike="sngStrike"` : "";
  const fillColor = opts?.color ?? "000000";
  const latin = opts?.typeface ? `<a:latin typeface="${escapeXmlAttribute(opts.typeface)}"/>` : "";
  const algn = opts?.align ? ` algn="${opts.align}"` : "";
  const anchor = opts?.anchor ?? "ctr";
  const wrap = opts?.wrap ? ` wrap="${opts.wrap}"` : "";
  const lnSpc = opts?.lineSpacingPts
    ? `<a:lnSpc><a:spcPts val="${opts.lineSpacingPts}"/></a:lnSpc>`
    : opts?.lineSpacing
      ? `<a:lnSpc><a:spcPct val="${opts.lineSpacing}"/></a:lnSpc>`
      : "";
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
        ${latin}
      </a:rPr>
      <a:t>${escapeXmlText(text)}</a:t>
    </a:r>
  </a:p>
</p:txBody>`;
}

export function escapeXmlText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, "&quot;");
}

export function wrapSlideXml(spTreeContent: string, backgroundXml = ""): string {
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

export function slideRelsXml(
  extraRels: { id: string; type: string; target: string; targetMode?: string }[] = [],
): string {
  const extras = extraRels
    .map((r) => {
      const tm = r.targetMode ? ` TargetMode="${r.targetMode}"` : "";
      return `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${tm}/>`;
    })
    .join("\n  ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
  ${extras}
</Relationships>`;
}

export interface SlideData {
  xml: string;
  rels: string;
}

export interface PptxBuildOptions {
  slides: SlideData[];
  charts?: Map<string, string>;
  media?: Map<string, Buffer>;
  contentTypesExtra?: string[];
  slideMasterXml?: string;
  slideMasterRelsXml?: string;
  slideLayoutXml?: string;
  slideSize?: { cx: number; cy: number; type?: string };
  themeXml?: string;
  defaultTextStyleXml?: string;
}

export async function buildPptx(options: PptxBuildOptions): Promise<Buffer> {
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
  const sldSzCx = options.slideSize?.cx ?? SLIDE_W;
  const sldSzCy = options.slideSize?.cy ?? SLIDE_H;
  const sldSzType = options.slideSize?.type ?? "screen16x9";
  const defaultTextStyleSection = options.defaultTextStyleXml
    ? `<p:defaultTextStyle>${options.defaultTextStyleXml}</p:defaultTextStyle>`
    : "";
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:sldMasterIdLst><p:sldMasterId r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${sldIdLst}</p:sldIdLst>
  <p:sldSz cx="${sldSzCx}" cy="${sldSzCy}" type="${sldSzType}"/>
  ${defaultTextStyleSection}
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
  zip.file("ppt/slideLayouts/slideLayout1.xml", options.slideLayoutXml ?? slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", options.themeXml ?? theme1);

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

const FIXTURE_OUT_DIR = join(__dirname, "fixtures");

export function savePptx(buffer: Buffer, name: string): void {
  mkdirSync(FIXTURE_OUT_DIR, { recursive: true });
  const path = join(FIXTURE_OUT_DIR, name);
  writeFileSync(path, buffer);
  console.log(`  Created: ${path}`);
}
